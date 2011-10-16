Attribute VB_Name = "mod_IA"
'Flamebird MX
'CopyRight$ (C) 2003-2007 Flamebird Team
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
Dim procesando_info As Boolean

' for user declared constants
Public userConstList() As String
' for user declared vars
Public userTypeList() As String
' maintains the state of variables includes the project being edited
Public includesNodes As New Collection
' used in the getline function () to maintain the status code in # includes
Dim inComment As Boolean
Dim MainDir As String
' list of functions to display the AutoComplete list
Public functionList() As String
' list of user defines functions/processes that we're go to show in intellISense autocomplete
Public userFunctionList() As String
' list of variables to show the autocomplete list
Public varList() As String
' contains the parameters of the functions and processes identified by the user
Public parameters() As String
' list of macros and their values
Public macros As New Collection
Public macrosNames As New Collection
' counter of language functions
Public iFunctionCount As Long
' control the analyzing status to not repeat analyzings at the same time
Public analyzingSource As Boolean


Public Function nodeExists(Key As String) As Boolean

End Function


'*****************************************************************
'** create include-nodes without reading the files              **
'** using the buffer builded at the first reading of the file   **
'*****************************************************************
Public Sub makeProgramTree(ByVal Filename As String, Optional isInclude As Boolean)

    Filename = LCase$(Filename)
    Filename = replace$(Filename, "/", "\")
    
    On Error GoTo Termina
    
    Dim i As Variant
    Dim fatherNode As String
    Dim indice As Integer
    
    ' if it's the main prg, set to cero
    If Not isInclude Then
        frmProgramInspector.tv_program.Visible = False
        frmProgramInspector.tv_program.Nodes.Clear
        ' build the function array with the language functions
        Dim tempList() As String
        
        With Ini
            .Path = App.Path & "\Help\functions.lan"
            .EnumerateAllSections tempList(), iFunctionCount
        
            ReDim functionList(0) As String
            For i = 1 To iFunctionCount
                ReDim Preserve functionList(i) As String
                functionList(i) = tempList(i)
            Next i
        End With
        
        ' reset everything to cero
        ReDim varList(0) As String
        ReDim parameters(0) As String
        ReDim userTypeList(0) As String
        ReDim userConstList(0) As String
        ReDim userFunctionList(0) As String
    End If
    
    Dim nodito As staticNode
    
    For Each nodito In includesNodes
    
        If LCase$(nodito.Filename) = LCase$(Filename) Then
            
            If nodito.varType <> "INCLUDE" Then
                fatherNode = nodito.father
                
                If frmProgramInspector.tv_program.Nodes.Exists(nodito.Key) = False Then

                    If fatherNode <> "" Then
                        'If nodito.varType <> "struct" Then
                            frmProgramInspector.tv_program.Nodes.Add frmProgramInspector.tv_program.Nodes.item(fatherNode), etvwChild, nodito.Key, nodito.name, nodito.Icon
                        'End If
                    Else
                        If (nodito.varAmbient = "const" And PI_ShowConsts) _
                            Or (nodito.varAmbient = "global" And PI_ShowGlobals) _
                            Or (nodito.varAmbient = "local" And PI_ShowLocals) _
                            Or (nodito.varAmbient = "private" And PI_ShowPrivates) _
                            Or nodito.varType = "function" _
                            Or nodito.varType = "process" _
                            Or nodito.varType = "struct" _
                        Then
                            Call frmProgramInspector.tv_program.Nodes.Add(, , nodito.Key, nodito.name, nodito.Icon)
                        End If
                    End If
                    
                    If (nodito.varType = "var" And nodito.father = "") Or nodito.varType = "type" Or nodito.varType = "struct" Then
                        ' add user-variable name for the code list
                        If nodito.varType <> "private" And nodito.varAmbient <> "const" Then
                            ReDim Preserve varList(UBound(varList) + 1) As String
                            varList(UBound(varList)) = nodito.name
                        End If
                    End If
                                            
                    If nodito.varAmbient = "const" And nodito.father = "" Then
                        ' add user-constant name for the code list
                        ReDim Preserve userConstList(UBound(userConstList) + 1) As String
                        userConstList(UBound(userConstList)) = nodito.name
                    End If
                    
                    If isInclude And (nodito.varType = "type" Or nodito.varType = "struct") Then
                        ' add user-struct name for the code list
                        ReDim Preserve userTypeList(UBound(userTypeList) + 1) As String
                        userTypeList(UBound(userTypeList)) = LCase$(nodito.name)
                    End If
                    
                    If nodito.varType = "function" Or nodito.varType = "process" Then
                        ' add user-function name for the code list
                        ReDim Preserve userFunctionList(UBound(userFunctionList) + 1) As String
                        userFunctionList(UBound(userFunctionList)) = nodito.name
                        
                        ' take it's params for the help tip
                        ReDim Preserve parameters(UBound(parameters) + 1) As String
                        parameters(UBound(parameters)) = nodito.parameters
                    End If
                End If
            Else
                makeProgramTree nodito.name, True
            End If
        End If
    Next nodito
    
Termina:
    
    If Not isInclude Then
        frmProgramInspector.tv_program.Visible = True
    End If


End Sub
' **************************************************************
' *** Returns true if the line is in const declaration zone ****
' **************************************************************
Public Function inConstDeclarationZone(lineaNum As Integer) As Boolean

    On Error Resume Next
    
    Dim frmPRG As frmDoc
    Set frmPRG = frmMain.ActiveForm
    Dim num As Integer
    Dim num2 As Integer
    Dim lineNum As Integer
    Dim palabra As String
    Dim linea As String
    Dim endsCount As Integer
    Dim importantLine As Boolean
    Dim nextLineCommented As Boolean
    
    lineNum = lineaNum
    inConstDeclarationZone = False
    
    While lineNum >= 0
        linea = frmPRG.cs.getLine(lineNum)
           
        importantLine = False
        
        '*******************************************
        '*********** Line optimization *************
        '*******************************************
            
        ' replace$ not visible chars with spaces
        linea = replace$(linea, vbTab, " ")
        linea = replace$(linea, vbNewLine, " ")
        linea = replace$(linea, vbCrLf, " ")
        linea = replace$(linea, vbCr, " ")
        linea = replace$(linea, vbLf, " ")
        linea = replace$(linea, vbNullChar, " ")
        linea = replace$(linea, vbBack, " ")
        linea = replace$(linea, vbFormFeed, " ")
        linea = replace$(linea, vbVerticalTab, " ")
        
        ' delete spaces
        linea = Trim$(linea)
        
        num = 0
        
        If inComment = False Then
            num = InStr(linea, "//")
            
            If num > 0 Then
                linea = Mid$(linea, 1, num - 1)
            End If
        End If
        
        If inComment = True Then
            num = InStr(linea, "/*")
            num2 = InStr(linea, "*/")
            
            If ((num > 0 And num2 < num) Or num = 0) And num2 > 0 Then
                linea = Mid$(linea, num2 + 2)
                inComment = False
            End If
        End If
        
        If inComment = False Then
            num = InStr(linea, "/*")
            While (num > 0)
                num2 = InStr(linea, "*/")
                If num2 > num Then
                    inComment = False
                    linea = Mid$(linea, 1, num - 1) & Mid$(linea, num2 + 2)
                Else
                    If inComment = False Then
                        linea = Mid$(linea, 1, num - 1)
                        inComment = True
                        nextLineCommented = True
                    End If
                End If
                If inComment = False Then
                    num = InStr(linea, "/*")
                Else
                    num = 0
                End If
            Wend
        End If
        
        
        If nextLineCommented = False Then
            If inComment = True Then
                linea = ""
            End If
        Else
            nextLineCommented = False
        End If
        '*******************************************
        '*********** code analization **************
        '*******************************************
        
        While Len(linea) > 0
            palabra = getWordRev(linea)
                    
            Select Case LCase$(palabra)
                ' FALSE cases
                Case "end":
                    inConstDeclarationZone = False
                    Exit Function
                Case "begin":
                    inConstDeclarationZone = False
                    Exit Function
                Case "process":
                    If lineNum = lineaNum Then
                        inConstDeclarationZone = True
                    Else
                        inConstDeclarationZone = False
                    End If
                    Exit Function
                Case "function":
                    If lineNum = lineaNum Then
                        inConstDeclarationZone = True
                    Else
                        inConstDeclarationZone = False
                    End If
                    Exit Function
                Case "program":
                    inConstDeclarationZone = False
                    Exit Function
                Case "local":
                    inConstDeclarationZone = False
                    Exit Function
                Case "private":
                    inConstDeclarationZone = False
                    Exit Function
                Case "global":
                    inConstDeclarationZone = False
                    Exit Function
                Case "struct":
                    inConstDeclarationZone = False
                    Exit Function
                Case "type":
                    inConstDeclarationZone = False
                    Exit Function
                '  TRUE cases
                Case "const":
                    inConstDeclarationZone = True
                    Exit Function
            End Select
            
        Wend
        
        
        lineNum = lineNum - 1
        
    Wend
End Function
' **************************************************************
' *** Returns true if the line is in a declaration zone ********
' **************************************************************
Public Function inDeclarationZone(lineaNum As Integer) As Boolean

    On Error Resume Next
    
    Dim frmPRG As frmDoc
    Set frmPRG = frmMain.ActiveForm
    Dim num As Integer
    Dim num2 As Integer
    Dim lineNum As Integer
    Dim palabra As String
    Dim linea As String
    Dim endsCount As Integer
    Dim importantLine As Boolean
    Dim nextLineCommented As Boolean
    
    lineNum = lineaNum
    inDeclarationZone = False
    
    While lineNum >= 0
        linea = frmPRG.cs.getLine(lineNum)
           
        importantLine = False
        
        '*******************************************
        '*********** Line optimization *************
        '*******************************************
            
        ' replace$ not visible chars with spaces
        linea = replace$(linea, vbTab, " ")
        linea = replace$(linea, vbNewLine, " ")
        linea = replace$(linea, vbCrLf, " ")
        linea = replace$(linea, vbCr, " ")
        linea = replace$(linea, vbLf, " ")
        linea = replace$(linea, vbNullChar, " ")
        linea = replace$(linea, vbBack, " ")
        linea = replace$(linea, vbFormFeed, " ")
        linea = replace$(linea, vbVerticalTab, " ")
        
        ' delete spaces
        linea = Trim$(linea)
        
        num = 0
        
        If inComment = False Then
            num = InStr(linea, "//")
            
            If num > 0 Then
                linea = Mid$(linea, 1, num - 1)
            End If
        End If
        
        If inComment = True Then
            num = InStr(linea, "/*")
            num2 = InStr(linea, "*/")
            
            If ((num > 0 And num2 < num) Or num = 0) And num2 > 0 Then
                linea = Mid$(linea, num2 + 2)
                inComment = False
            End If
        End If
        
        If inComment = False Then
            num = InStr(linea, "/*")
            While (num > 0)
                num2 = InStr(linea, "*/")
                If num2 > num Then
                    inComment = False
                    linea = Mid$(linea, 1, num - 1) & Mid$(linea, num2 + 2)
                Else
                    If inComment = False Then
                        linea = Mid$(linea, 1, num - 1)
                        inComment = True
                        nextLineCommented = True
                    End If
                End If
                If inComment = False Then
                    num = InStr(linea, "/*")
                Else
                    num = 0
                End If
            Wend
        End If
        
        
        If nextLineCommented = False Then
            If inComment = True Then
                linea = ""
            End If
        Else
            nextLineCommented = False
        End If
        '*******************************************
        '*********** code analization **************
        '*******************************************
        
        While Len(linea) > 0
            palabra = getWordRev(linea)
                    
            Select Case LCase$(palabra)
                ' FALSE cases
                Case "end":
                    inDeclarationZone = False
                    Exit Function
                Case "begin":
                    inDeclarationZone = False
                    Exit Function
                Case "process":
                    If lineNum = lineaNum Then
                        inDeclarationZone = True
                    Else
                        inDeclarationZone = False
                    End If
                    Exit Function
                Case "function":
                    If lineNum = lineaNum Then
                        inDeclarationZone = True
                    Else
                        inDeclarationZone = False
                    End If
                    Exit Function
                Case "program":
                    inDeclarationZone = False
                    Exit Function
                '  TRUE cases
                Case "local":
                    inDeclarationZone = True
                    Exit Function
                Case "private":
                    inDeclarationZone = True
                    Exit Function
                Case "const":
                    inDeclarationZone = True
                    Exit Function
                Case "global":
                    inDeclarationZone = True
                    Exit Function
                Case "struct":
                    inDeclarationZone = True
                    Exit Function
                Case "type":
                    inDeclarationZone = True
                    Exit Function
            End Select
            
        Wend
        
        
        lineNum = lineNum - 1
        
    Wend
End Function

' **************************************************************
'    takes the parameter list of a function or a process
' **************************************************************

Private Function getParameters(linea As String)
    On Error Resume Next
    Dim pStart As Integer
    Dim pEnd As Integer
    Dim strResult As String
    
    strResult = " "
    
    pStart = InStr(linea, "(")
    If pStart > 0 Then
        pEnd = InStr(pStart, linea, ")")
        If pEnd > 0 Then
           strResult = Mid$(linea, pStart + 1, pEnd - 1 - pStart)
        End If
    End If
    
    getParameters = strResult
End Function

Public Function existTreeForFile(ByVal Filename) As Boolean
    Filename = LCase$(Filename)
    Filename = replace$(Filename, "/", "\")
    Dim nodito As staticNode
    
    For Each nodito In includesNodes
        If LCase$(replace$(nodito.Filename, "/", "\")) = LCase$(Filename) Then
            existTreeForFile = True
            Exit Function
        End If
    Next nodito

End Function

' **************************************************************
'        Builds the declaration-tree of fucntions of the prg
'        now only in buffer
' **************************************************************

Public Sub MakeProgramIndex(ByVal Filename As String, Optional isInclude As Boolean)

    On Error GoTo Termina
    
    If analyzingSource Or Not PI_Active Then Exit Sub
    
    'If there is an open project with defined mainsource
    'Directly sent to make that file is only logical, Right$?

    If Not openedProject Is Nothing And isInclude = False Then
        If openedProject.FileExist(Filename) Then
            If openedProject.mainSource <> "" Then
                Filename = makePathForProject(openedProject.mainSource)
            End If
        End If
    End If
    
    
    Filename = LCase$(Filename)
    Filename = replace$(Filename, "/", "\")
    
    Dim nodito As staticNode
    Dim formulario As frmDoc
     
     
    ' if this is an include and already exist nodes of this file, send to build from the buffer
    If isInclude Then
        If existTreeForFile(Filename) Then
                ' look if exists the form of this include
                ' and if it is needed, refresh to not send to back
                Set formulario = FindFileForm(Filename)
                
                Dim refrescar As Boolean
                refrescar = False
                If Not formulario Is Nothing Then
                    If formulario.mustRefresh = True Then
                        refrescar = True
                    End If
                End If
                
                If refrescar = False Then
                    makeProgramTree Filename, True
                    Exit Sub
                End If
        End If
    Else
        
        ' Insteadif it is not include
        ' Means that they sent him to refresh or build the tree
        ' But ... let's see if it already exists in the buffer, to no wait a lot
        
        Set formulario = FindFileForm(Filename)
        
        If Not formulario Is Nothing Then
            
            Dim imMainPrg As Boolean
            imMainPrg = False
            
            ' Let's see if was send to refresh the main PRG
            ' To consider whether you need to refresh some of the includes
            If Not openedProject Is Nothing Then
                If openedProject.mainSource <> "" Then
                    Dim mainPath As String
                    mainPath = makePathForProject(openedProject.mainSource)
                    mainPath = LCase$(mainPath)
                    mainPath = replace$(mainPath, "/", "\")
                    If Filename = mainPath Then
                        imMainPrg = True
                    End If
                End If
            End If
    
            If imMainPrg Then
                Dim item As Variant
                For Each item In openedProject.Files
                    Dim include As String
                    include = CStr(item)
                    Dim formu As frmDoc
                    Set formu = FindFileForm(makePathForProject(include))
                    If Not formu Is Nothing Then
                        If formu.mustRefresh Then
                            formulario.mustRefresh = True
                            Exit For
                        End If
                    End If
                Next
            End If
    
            ' this is unjust, cause can be a include that needs to be refreshed
            If formulario.mustRefresh = False Then
                If existTreeForFile(Filename) Then
                    makeProgramTree Filename
                    Exit Sub
                End If
            End If
        End If
        
        frmMain.StatusBar.PanelText("MAIN") = "Collecting info about the project"
    End If
    
    Dim srcFile As New cReadFile ' class that reads the file
    Dim linea As String
    Dim palabra As String
    Dim fatherNode As String
    Dim fatherType As String
    Dim declarationType As String
    Dim varType As String
    Dim waitFor As String
    Dim lineNum As Variant
    Dim fileNum As String
    Dim endsCount As Integer
    Dim imagen As Integer
    Dim returnValue As String
    Dim num As Variant, num2 As Long
    Dim structFather As String
    Dim structFatherType As String
    Dim nextLineCommented As Boolean
    Dim lineaTemp As String
    Dim nextWord As String
    Dim tempNode As Variant
    Dim fixNode As staticNode
    Dim fixExistsNode As Boolean
    
    Dim newKey As String
    
    ' if it is the main file, set to zero
    If Not isInclude Then
        inComment = False
        MainDir = FSO.GetParentFolderName(Filename) & "\"
        ' empty the user Type List
        ReDim userTypeList(0) As String
        ' clear the macros
        Set macros = New Collection
        Set macrosNames = New Collection
    End If
    
    For Each nodito In includesNodes
        If LCase$(nodito.Filename) = LCase$(Filename) Then
            includesNodes.Remove nodito.Key
        End If
    Next nodito
    
    srcFile.Filename = Filename
    'Screen.MousePointer = vbHourglass
    analyzingSource = True
    
    ' for each line in the PRG
    While srcFile.canRead
        
        If lineNum Mod 5 = 0 Then
            DoEvents
            frmMain.StatusBar.PanelText("MAIN") = "analyzing file structure: " & CLng(lineNum * 100 / formulario.cs.LineCount) & "% done... Please wait"
        End If
        
        ' get a line
        linea = srcFile.getLine
        
        lineNum = lineNum + 1
        
        '*******************************************
        '*********** Line clearing            ******
        '*******************************************
            
        ' replace$ non visible chars with spaces
        linea = replace$(linea, vbTab, " ")
        linea = replace$(linea, vbNewLine, " ")
        linea = replace$(linea, vbCrLf, " ")
        linea = replace$(linea, vbCr, " ")
        linea = replace$(linea, vbLf, " ")
        linea = replace$(linea, vbNullChar, " ")
        linea = replace$(linea, vbBack, " ")
        linea = replace$(linea, vbFormFeed, " ")
        linea = replace$(linea, vbVerticalTab, " ")
        
        ' clear the spaces
        linea = Trim$(linea)
        
        ' TODO analyze the line and replace$ all macros
    '    Dim macro_value As String
    '    For Each macro_value In macros
    '    Do
    '    while InStr$$(linea,
    '    Next
        
        num = 0
        num2 = 0
        
       
        If inComment = False Then
            num = InStr(linea, "//")
            
            If num > 0 Then
                linea = Mid$(linea, 1, num - 1)
            End If
        End If
        
        If inComment = True Then
            num = InStr(linea, "/*")
            num2 = InStr(linea, "*/")
            
            If ((num > 0 And num2 < num) Or num = 0) And num2 > 0 Then
                linea = Mid$(linea, num2 + 2)
                inComment = False
            End If
        End If
        
        If inComment = False Then
            num = InStr(linea, "/*")
            While (num > 0)
                num2 = InStr(linea, "*/")
                If num2 > num Then
                    inComment = False
                    linea = Mid$(linea, 1, num - 1) & Mid$(linea, num2 + 2)
                Else
                    If inComment = False Then
                        linea = Mid$(linea, 1, num - 1)
                        inComment = True
                        nextLineCommented = True
                    End If
                End If
                If inComment = False Then
                    num = InStr(linea, "/*")
                Else
                    num = 0
                End If
            Wend
        End If
        
        
        If nextLineCommented = False Then
            If inComment = True Then
                linea = ""
            End If
        Else
            nextLineCommented = False
        End If
        '*******************************************
        '*********** CODE ANALYSIS      ************
        '*******************************************
        
        While Len(linea) > 0
            On Error GoTo proximaVuelta
            palabra = getWord(linea)
            
            ' is the word a macro?
            Dim macro_value As Variant
            Dim macindex As Integer
            Dim macroname As Variant
            
            While macindex < macros.count
                macro_value = macros.item(macindex + 1)
                macroname = macrosNames.item(macindex + 1)
                If LCase$(palabra) = Trim$(LCase$(macroname)) Then
                    palabra = ""
                    linea = macro_value & " " & linea
                End If
                macindex = macindex + 1
            Wend
            
            If palabra <> "" Then
                If endsCount > 0 Then
                    Select Case LCase$(palabra)
                        'Case "#if": endsCount = endsCount + 1
                        'Case "#ifdef": endsCount = endsCount + 1
                        'Case "#ifndef": endsCount = endsCount + 1
                        Case "if": endsCount = endsCount + 1
                        Case "for": endsCount = endsCount + 1
                        Case "case": endsCount = endsCount + 1
                        Case "switch": endsCount = endsCount + 1
                        Case "end": endsCount = endsCount - 1
                        'Case "#endif": endsCount = endsCount - 1
                        Case "clone": endsCount = endsCount + 1
                        Case "while": endsCount = endsCount + 1
                        Case "loop": endsCount = endsCount + 1
                        Case "from": endsCount = endsCount + 1
                    End Select
                Else
                
                    If isUserDefinedType(palabra) Or isDefinedType(palabra) Then
                        
                        If declarationType = "" Or declarationType = "variables" Then
                            ' copy of line to take the next word
                            lineaTemp = linea
                            ' this must be the name
                            nextWord = getWord(lineaTemp)
                            ' the next word must be (
                            nextWord = getWord(lineaTemp)
                            
                            If nextWord = "(" Then
                                ' if there are parameters to be defined, then
                                ' it's not a variable, it's a function
                                declarationType = "functionName"
                            End If
                        End If
                    End If
                    
                    Dim Inicio As String
                    Dim Largo As String
                    
                            
                    Select Case LCase$(palabra)
                        Case "#define":
                            
                            Dim macroValue As String
                            Dim Macro As String
                            
                            Largo = InStr(linea, " ")
                            If Largo > 0 Then
                                ' can have value after the space
                                Macro = Trim$(Left$(linea, Largo))
                                macroValue = Trim$(Mid$(linea, Largo))
                            Else
                                ' there's no value, only a name
                                Macro = Trim$(linea)
                            End If
                            
                            If macroValue <> "" Then
                            On Error GoTo macroerror
                                macros.Add macroValue, Macro
macroerror:                     macrosNames.Add Macro, Macro
                            End If
                            
                            
                            'On Error GoTo Esquiva1
    
                            'ReDim Preserve userTypeList(UBound(userTypeList) + 1) As String
                            'userTypeList(UBound(userTypeList)) = LCase$(Macro)
    
                            linea = ""
                            declarationType = ""
                            
                        Case "include":
                            
                            Dim incluir As String
                            Inicio = InStr(linea, Chr(34)) + 1
                            Largo = InStrRev(linea, Chr(34)) - Inicio
                            incluir = Mid$(linea, Inicio, Largo)
                            
                            If InStr(linea, ";") <> 0 Then
                                linea = Mid$(linea, InStr(linea, ";") + 1)
                            End If
                            
                            ' here we see if this is relative dir
                            If FSO.GetDriveName(incluir) = "" Then
                                ' so, add to the main project dir
                               incluir = MainDir & incluir
                            End If
                            
                            ' if exist
                            If FSO.FileExists(incluir) Then
                            
                               incluir = LCase$(incluir)
                               
                               
                               'On Error GoTo Esquiva1
                               
                               newKey = "INCLUDE" & "|" & incluir & "|" & Filename
                                                                                     
                               Set tempNode = New staticNode
                               includesNodes.Add tempNode, newKey
                            
                               tempNode.Filename = Filename
                               tempNode.name = incluir
                               tempNode.Key = newKey
                               tempNode.varType = "INCLUDE"
                            
                               Call MakeProgramIndex(incluir, True)
                                
Esquiva1:
    'On Error GoTo Termina
                            End If
                            
                            declarationType = ""
                        Case "import": declarationType = "importName"
                        Case "process": declarationType = "processName"
                            If G_ProcHelpLine = 0 Then
                                formulario.cs.SetDivider lineNum - 1, True
                            ElseIf G_ProcHelpLine = 1 Then
                                formulario.cs.SetDivider lineNum - 2, True
                            End If
                        Case "function": declarationType = "functionName"
                            If G_ProcHelpLine = 0 Then
                                formulario.cs.SetDivider lineNum - 1, True
                            ElseIf G_ProcHelpLine = 1 Then
                                formulario.cs.SetDivider lineNum - 2, True
                            End If
                        Case "const":
                            declarationType = "variables"
                            varType = "const"
                            fatherNode = ""
                            fatherType = ""
                        Case "private":
                            declarationType = "variables"
                            varType = "private"
                        Case "global":
                            declarationType = "variables"
                            varType = "global"
                            fatherNode = ""
                            fatherType = ""
                        Case "local":
                            declarationType = "variables"
                            varType = "local"
                            If endsCount = 0 Then
                                fatherNode = ""
                                fatherType = ""
                            End If
                        Case "end":
                            If varType = "struct" Then
                                declarationType = "variables"
                                varType = Left$(fatherNode, InStr(fatherNode, "|") - 1)
                                fatherType = structFatherType
                                fatherNode = structFather
                            Else
                                declarationType = ""
                                varType = ""
                                fatherNode = ""
                                fatherType = ""
                            End If
                        Case "#define"
                            'declarationType = "macroDefinition"
                        Case "#if"
                            declarationType = ""
                            'endsCount = 1
                        Case "struct"
                            declarationType = "struct"
                            endsCount = 0
                        Case "type"
                            declarationType = "type"
                            varType = "type"
                        Case "begin":
                            declarationType = ""
                            endsCount = 1
                        Case Chr(34):
                            If declarationType = "variables" And waitFor = "" Then
                                waitFor = Chr(34) ' wait until this word
                            End If
                        Case "=":
                            If declarationType = "variables" And waitFor = "" Then
                                waitFor = ";" ' wait until this word
                            End If
                        Case "[":
                            If declarationType = "variables" And waitFor = "" Then
                                waitFor = "]" ' wait until this word
                            End If
                        
                        Case Else:
                            If waitFor <> "" Then ' if we were waiting a word
                                If palabra = waitFor Then ' if we find what we wait for
                                   waitFor = "" ' continue with normal declaration
                                End If
                            Else
                                Select Case declarationType
                               ' Case "macroDefinition":
                                    ' fixed in 2006 revision
                                    ' a macro is NOT a data type
                                    
                                   ' ReDim Preserve userTypeList(UBound(userTypeList) + 1) As String
                                   ' userTypeList(UBound(userTypeList)) = LCase$(palabra)
                                    
                                   ' declarationType = ""
                                    
                                Case "type":
                                    
                                    newKey = "type" & "|" & palabra & "|" & Filename
                                    
                                    'On Error GoTo Esquiva2
                                    
                                    
                                    ' add declaration to the main buffer
                                    
                                    Set tempNode = New staticNode
                                    includesNodes.Add tempNode, newKey
                                    
                                    tempNode.father = ""
                                    tempNode.Filename = Filename
                                    tempNode.Icon = 8 - 1
                                    tempNode.Key = newKey
                                    tempNode.name = palabra
                                    tempNode.varType = "type"
                                    
                                    ' add data type to main list
                                    ReDim Preserve userTypeList(UBound(userTypeList) + 1) As String
                                    userTypeList(UBound(userTypeList)) = LCase$(palabra)
                                    
        
                                    declarationType = "variables"
        
                                    fatherNode = newKey
                                    fatherType = "type"
                                    
Esquiva2:
    'On Error GoTo Termina
                                Case "struct":
                                    
                                    newKey = varType & "|" & palabra & "|" & fatherNode & "|" & Filename
                                    
                                    'On Error GoTo Esquiva3
                                    
                                    Set tempNode = New staticNode
                                    includesNodes.Add tempNode, newKey
                                    
                                    structFatherType = fatherType
                                    structFather = fatherNode
                                    
                                    
                                    tempNode.father = fatherNode
                                    tempNode.Filename = Filename
                                    tempNode.Icon = 7 - 1
                                    tempNode.Key = newKey
                                    tempNode.name = palabra
                                    tempNode.varType = "struct"
                                    
                                                                    
                                    declarationType = "variables"
                                    fatherNode = newKey
                                    fatherType = "struct"
                                    
                                    varType = "struct"
                                    
Esquiva3:
    'On Error GoTo Termina
            
                                Case "functionName":
                                    
                                    
                                    ' if it's a type returning function ( int getValue(); )
                                    If isUserDefinedType(palabra) Or isDefinedType(palabra) Then
                                        ' copy of line to capture the next word
                                        lineaTemp = linea
                                        ' this must be the name
                                        nextWord = getWord(lineaTemp)
                                        newKey = "function" & "|" & nextWord & "|" & Filename
                                    Else
                                        ' if it's a common function declare ( Function gameMain(); )
                                        newKey = "function" & "|" & palabra & "|" & Filename
                                    End If
                                    
                                   ' On Error GoTo Esquiva4
                                    'TEMPORAL BUGFIX:
                                    ' If are used declares for functions, they appear twice
                                    fixExistsNode = False
                                    For Each fixNode In includesNodes
                                        If fixNode.name = palabra Then
                                            fixExistsNode = True
                                            Exit For
                                        End If
                                    Next
                                    If fixExistsNode = False Then
                                    Set tempNode = New staticNode
                                        includesNodes.Add tempNode, newKey
                                        
                                        tempNode.father = ""
                                        tempNode.Filename = Filename
                                        tempNode.Icon = 9 - 1
                                        tempNode.Key = newKey
                                        If isUserDefinedType(palabra) Or isDefinedType(palabra) Then
                                            tempNode.name = nextWord
                                        Else
                                            tempNode.name = palabra
                                        End If
                                        tempNode.varType = "function"
                                        tempNode.parameters = getParameters(linea)
                
                                        declarationType = ""
                
                                        fatherNode = newKey
                                        fatherType = "function"
                                        tempNode.lineNum = lineNum
                                    End If
                                    
Esquiva4:
    'On Error GoTo Termina
                                Case "importName":
                                    
                                    newKey = "import" & "|" & palabra & "|" & Filename
                                    
                                    'On Error GoTo Esquiva5
                                    
                                    Set tempNode = New staticNode
                                    includesNodes.Add tempNode, newKey
                                    
                                    tempNode.father = ""
                                    tempNode.Filename = Filename
                                    tempNode.Icon = 19 - 1
                                    tempNode.Key = newKey
                                    tempNode.name = palabra
                                    tempNode.varType = "import"
                                    
                                    
                                    declarationType = ""
                                    
Esquiva5:
    'On Error GoTo Termina
                                
                                Case "processName":
                                    
                                    newKey = "process" & "|" & palabra & "|" & Filename
                                    
                                    'On Error GoTo Esquiva6
                                    'TEMPORAL BUGFIX:
                                    ' If are used declares for processes, they appear twice
                                    fixExistsNode = False
                                    For Each fixNode In includesNodes
                                        If fixNode.name = palabra Then
                                            fixExistsNode = True
                                            Exit For
                                        End If
                                    Next
                                    
                                    If fixExistsNode = False Then
                                        Set tempNode = New staticNode
                                        
                                        includesNodes.Add tempNode, newKey
                                        
                                        tempNode.father = ""
                                        tempNode.Filename = Filename
                                        tempNode.Icon = 10 - 1
                                        tempNode.Key = newKey
                                        tempNode.name = palabra
                                        tempNode.varType = "process"
                                        tempNode.parameters = getParameters(linea)
                
                                        declarationType = ""
                
                                        fatherNode = newKey
                                        fatherType = "process"
                                        tempNode.lineNum = lineNum
                                    End If
                                    
Esquiva6:
    'On Error GoTo Termina
                                    
                                Case "variables":
                                    ' if the word isn't part of the language
                                    If isUserDefinedType(palabra) = False And isDefinedType(palabra) = False And isOperator(palabra) = False Then
                                        
                                        Select Case varType
                                            Case "const":   imagen = 1 - 1
                                            Case "global":  imagen = 2 - 1
                                            Case "local":   imagen = 3 - 1
                                            Case "private": imagen = 4 - 1
                                            Case "public":  imagen = 5 - 1
                                            Case "table":   imagen = 6 - 1
                                            Case "struct":  imagen = 7 - 1
                                            Case "type":    imagen = 8 - 1
                                            Case Else:      imagen = -1
                                        End Select
                                        
                                        newKey = lineNum & "|" & imagen & "|" & palabra & "|" & fatherNode & "|" & Filename
                                        
                                        'On Error GoTo Esquiva7
                                        
                                        Set tempNode = New staticNode
                                        includesNodes.Add tempNode, newKey
                                    
                                        tempNode.father = fatherNode
                                        tempNode.Filename = Filename
                                        tempNode.Icon = imagen
                                        tempNode.Key = newKey
                                        tempNode.name = palabra
                                        tempNode.varAmbient = varType
                                            tempNode.varType = "var"
                                           ' MsgBox tempNode.name
                                        'Else
                                            'MsgBox tempNode.name
'                                        End If
                                        Debug.Print tempNode.name
                                        If tempNode.name = "s_rune" Then
                                            Debug.Print "Achtung!"
                                        End If
Esquiva7:
    'On Error GoTo Termina
        
                                    End If
            
                                End Select
                            End If
                    End Select
                End If 'if endscount
            End If
proximaVuelta:
        Wend
    Wend
    
    'Screen.MousePointer = vbDefault
    analyzingSource = False
    
Termina:
    If Not isInclude Then
        frmMain.StatusBar.PanelText("MAIN") = ""
        makeProgramTree Filename
    End If

End Sub

' **************************************************************
'          get only one word from the line
' **************************************************************


Public Function getWord(ByRef line As String) As String
    Dim num() As Variant
    Dim num2 As Variant
    Dim num3 As Variant
    Dim operatorsList() As String
    Dim returnValue As String
    
    ' list of operators used in declaration
    ReDim operatorsList(LBound(operatorList) To (UBound(operatorList) + 1))
    
    Dim i As Long
    
    For i = LBound(operatorList) To UBound(operatorList)
      operatorsList(i) = operatorList(i)
    Next
    operatorsList(UBound(operatorsList)) = " "
    
    ReDim num(LBound(operatorsList) To UBound(operatorsList))
    
    getWord = ""
    
    For i = LBound(operatorsList) To UBound(operatorsList)
        num(i) = InStr(line, operatorsList(i))
    Next i
    
    num3 = 0
    num2 = Len(line)
    For i = LBound(num) To UBound(num)
        If num(i) > 0 And num(i) <= num2 Then
            num2 = num(i)
            num3 = num(i)
        End If
    Next i
            
    ' search the smallest num
    
    If num3 > 0 Then
        ' if long is 1, it's an operator
        If num3 = 1 Then
            returnValue = Trim$(Mid$(line, 1, num3))
            line = Trim$(Mid$(line, num3 + 1))
        Else
        ' if it's not a word, we delete the operator
            returnValue = Trim$(Mid$(line, 1, num3 - 1))
            line = Trim$(Mid$(line, num3))
        End If
    Else
    ' if there are no more word endings, the line is finished, so we return the word as line
        returnValue = Trim$(line)
        line = ""
    End If
            
    If InStr(returnValue, vbTab) = 1 Then
        returnValue = Mid$(returnValue, 2)
    Else
        If InStr(returnValue, vbTab) = Len(returnValue) Then
            If Len(returnValue) = 0 Then
                returnValue = ""
            Else
                returnValue = Mid$(returnValue, 1, Len(returnValue) - 1)
            End If
        End If
    End If
            
    getWord = returnValue
End Function

' **************************************************************
'       Gets only one word from de line
'       From Right$ to Left$ (reverse)
' **************************************************************


Public Function getWordRev(ByRef linea As String) As String
    Dim num() As Long
    Dim num2 As Variant
    Dim num3 As Variant
    Dim operatorsList() As String
    Dim returnValue As String
    Dim i As Integer
    
    ' delete not useful chars
    linea = Trim$(replace$(linea, vbTab, " "))
            
    ' list of operators used in declaration
    ReDim operatorsList(LBound(operatorList) To (UBound(operatorList) + 1))
    
    For i = LBound(operatorList) To UBound(operatorList)
      operatorsList(i) = operatorList(i)
    Next
    operatorsList(UBound(operatorsList)) = " "
    
    ReDim num(LBound(operatorsList) To UBound(operatorsList))
    
    getWordRev = ""
    
    i = 0
    
    For i = LBound(operatorsList) To UBound(operatorsList)
        num(i) = InStrRev(linea, operatorsList(i))
    Next i
    
    ' search the biggest num
    num3 = 0
    num2 = 0
    For i = LBound(num) To UBound(num)
        If num(i) > num2 Then
            num2 = num(i)
            num3 = num(i)
        End If
    Next i
    
    
    If num3 > 0 Then
        ' if the case is as long as the line, then it's an operator
        If num3 = Len(linea) Then
            returnValue = Right$(linea, 1)
            linea = Trim$(Mid$(linea, 1, num3 - 1))
        Else
            returnValue = Trim$(Mid$(linea, num3 + 1))
            linea = Trim$(Mid$(linea, 1, num3))
        End If
    Else
    ' if there are no more word endings, the line is finished, so we return the word as line
        returnValue = Trim$(linea)
        linea = ""
    End If
              
    getWordRev = returnValue
    
End Function

Public Function isUserDefinedType(ByVal palabra As String) As Boolean
    isUserDefinedType = False
    ' if it's not an user defined type
    If Not IsEmpty(LBound(userTypeList)) Then
        Dim i As Long
        For i = LBound(userTypeList) To UBound(userTypeList)
            If LCase$(palabra) = LCase$(userTypeList(i)) Then
                isUserDefinedType = True
                Exit Function
            End If
        Next i
    End If
End Function

Public Function isDefinedType(ByVal palabra As String) As Boolean
    Dim num As Double
    
    isDefinedType = False
    
    For num = LBound(typeList) To UBound(typeList)
        If LCase$(palabra) = LCase$(typeList(num)) Then
            isDefinedType = True
            Exit Function
        End If
    Next num
End Function

' looks in fdl.lan for the operators by the giving word
Public Function isOperator(ByVal sword As String) As Boolean
Dim num As Double

isOperator = False

For num = LBound(operatorList) To UBound(operatorList)
    If LCase$(sword) = LCase$(operatorList(num)) Then
        isOperator = True
        Exit Function
    End If
Next num
End Function

' returns true if the word passed as parameter is language reserved word
Public Function isReservedWord(ByVal sword As String) As Boolean
    Dim line As String
    Dim num As Double
    
    num = FreeFile()
    
    isReservedWord = False
    If isUserDefinedType(sword) Then
        isReservedWord = True
        Exit Function
    End If
    Open App.Path & "\Help\fdl.lan" For Input As #num
        Do Until EOF(num)
            ' reads a line
            Line Input #num, line
            ' jumps comments
            If InStr(line, "//#") <> 1 And line <> "" Then
                If LCase$(sword) = LCase$(Trim$(line)) Then
                    isReservedWord = True
                    Close #num
                    Exit Function
                End If
            End If
        Loop
    Close #num

End Function
' returns true if the giving word is language reserved function
Public Function isReservedFunction(ByVal sword As String) As Boolean
Dim num As Double

isReservedFunction = False

For num = LBound(functionList) To UBound(functionList)
    If LCase$(sword) = LCase$(functionList(num)) Then
        isReservedFunction = True
        Exit Function
    End If
Next num
End Function

' returns true if the giving word is user defined function
Public Function isUserDefinedFunction(ByVal sword As String) As Boolean
Dim num As Double

isUserDefinedFunction = False

For num = LBound(userFunctionList) To UBound(userFunctionList)
    If LCase$(sword) = LCase$(userFunctionList(num)) Then
        isUserDefinedFunction = True
        Exit Function
    End If
Next num
End Function

Public Function clearLine(ByVal sLine As String) As String

    ' replace$$s all not visible chars with spaces
    sLine = replace$(sLine, vbTab, " ")
    sLine = replace$(sLine, vbNewLine, " ")
    sLine = replace$(sLine, vbCrLf, " ")
    sLine = replace$(sLine, vbCr, " ")
    sLine = replace$(sLine, vbLf, " ")
    sLine = replace$(sLine, vbNullChar, " ")
    sLine = replace$(sLine, vbBack, " ")
    sLine = replace$(sLine, vbFormFeed, " ")
    sLine = replace$(sLine, vbVerticalTab, " ")
            
    clearLine = sLine
            
End Function
