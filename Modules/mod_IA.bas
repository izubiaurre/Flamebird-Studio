Attribute VB_Name = "mod_IA"
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
Dim procesando_info As Boolean
'para los tipos de datos declarados por el usuario
Public typeList() As String
'mantiene el estado de las variables de includes del proyecto que se esta editando
Public includesNodes As New Collection
'usadas en la funcion getLine() para mantener el estado del codigo entre #includes
Dim inComment As Boolean
Dim MainDir As String
'lista de funciones que va a mostrar la lista de autocompletar
Public functionList() As String
'lista de variables para mostrar en la lista de autocompletar
Public varList() As String
'contiene los parametros de las funciones y procesos declarados por el usuario
Public parameters() As String
'lista de macros y sus valores
Public macros As New Collection
Public macrosNames As New Collection
'contador de funciones propias del lenguaje
Public iFunctionCount As Long


Public Function nodeExists(Key As String) As Boolean

End Function


'*****************************************************************
'** crea los nodos de los includes sin leer de los archivos     **
'** usando un buffer creado la primera vez que se abre un archivo*
'*****************************************************************
Public Sub makeProgramTree(ByVal filename As String, Optional isInclude As Boolean)

filename = LCase(filename)
filename = replace(filename, "/", "\")

On Error GoTo Termina

Dim i As Variant
Dim fatherNode As String
Dim indice As Integer

' si se trata del archivo principal setea a cero
If Not isInclude Then
    frmProgramInspector.tv_program.Visible = False
    frmProgramInspector.tv_program.Nodes.Clear
    ' arma el array de funciones con las funciones basicas
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
    
    ' reseteamos todo a cero
    ReDim varList(0) As String
    ReDim parameters(0) As String
    ReDim typeList(0) As String
End If

Dim nodito As staticNode

For Each nodito In includesNodes
    If LCase(nodito.filename) = LCase(filename) Then
        
        If nodito.varType <> "INCLUDE" Then
            fatherNode = nodito.father
            
            If frmProgramInspector.tv_program.Nodes.Exists(nodito.Key) = False Then
            
                If fatherNode <> "" Then
                    Call frmProgramInspector.tv_program.Nodes.Add(frmProgramInspector.tv_program.Nodes.item(fatherNode), etvwChild, nodito.Key, nodito.name, nodito.Icon)
                Else
                    Call frmProgramInspector.tv_program.Nodes.Add(, , nodito.Key, nodito.name, nodito.Icon)
                End If
                
                If (nodito.varType = "variables" And nodito.father = "") Or nodito.varType = "type" Or nodito.varType = "struct" Then
                    ' agrega el nombre de la funcion recien creada a la lista para autocompletado
                    If nodito.varType <> "private" Then
                        ReDim Preserve varList(UBound(varList) + 1) As String
                        varList(UBound(varList)) = nodito.name
                    End If
                End If
                
                If isInclude And (nodito.varType = "type" Or nodito.varType = "struct") Then
                    ReDim Preserve typeList(UBound(typeList) + 1) As String
                    typeList(UBound(typeList)) = LCase(nodito.name)
                End If
                
                If nodito.varType = "function" Or nodito.varType = "process" Then
                    ' agrega el nombre de la funcion recien creada a la lista para autocompletado
                    ReDim Preserve functionList(UBound(functionList) + 1) As String
                    functionList(UBound(functionList)) = nodito.name
                    
                    ' toma sus parametros para mostrarlos en el tip
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
' *** Devuelve true si la linea en cuestion esta dentro de *****
' *** una zona de declaración. *********************************
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
    
    num = 0
    
    If inComment = False Then
        num = InStr(linea, "//")
        
        If num > 0 Then
            linea = Mid(linea, 1, num - 1)
        End If
    End If
    
    If inComment = True Then
        num = InStr(linea, "/*")
        num2 = InStr(linea, "*/")
        
        If ((num > 0 And num2 < num) Or num = 0) And num2 > 0 Then
            linea = Mid(linea, num2 + 2)
            inComment = False
        End If
    End If
    
    If inComment = False Then
        num = InStr(linea, "/*")
        While (num > 0)
            num2 = InStr(linea, "*/")
            If num2 > num Then
                inComment = False
                linea = Mid(linea, 1, num - 1) & Mid(linea, num2 + 2)
            Else
                If inComment = False Then
                    linea = Mid(linea, 1, num - 1)
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
    '*********** ANALISIS DE CODIGO ************
    '*******************************************
    
    While Len(linea) > 0
        palabra = getWordRev(linea)
                
        Select Case LCase(palabra)
            'casos FALSE
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
            'casos TRUE
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
'
'    EXTRAE LA LISTA DE PARAMETROS DE UNA FUNCION O PROCESO
'
' **************************************************************

Private Function getParameters(linea As String)
On Error Resume Next
Dim pStart As Integer
Dim pEnd As Integer
Dim resultado As String

resultado = " "

pStart = InStr(linea, "(")
If pStart > 0 Then
    pEnd = InStr(pStart, linea, ")")
    If pEnd > 0 Then
       resultado = Mid(linea, pStart + 1, pEnd - 1 - pStart)
    End If
End If

getParameters = resultado
End Function

Public Function existTreeForFile(ByVal filename) As Boolean
filename = LCase(filename)
filename = replace(filename, "/", "\")
Dim nodito As staticNode

For Each nodito In includesNodes
    If LCase(replace(nodito.filename, "/", "\")) = LCase(filename) Then
        existTreeForFile = True
        Exit Function
    End If
Next nodito

End Function

' **************************************************************
'
'        ARMA EL ARBOL CON LA DECLARACIONES DEL PROGRAMA
'         AHORA SOLO EN BUFFER
'
' **************************************************************

Public Sub MakeProgramIndex(ByVal filename As String, Optional isInclude As Boolean)

On Error GoTo Termina

' si hay un proyecto abierto con mainsource definido
' directamente mandamos a hacer ese achivo, es lo logico, no?
If Not openedProject Is Nothing And isInclude = False Then
    If openedProject.FileExist(filename) Then
        If openedProject.mainSource <> "" Then
            filename = makePathForProject(openedProject.mainSource)
        End If
    End If
End If


filename = LCase(filename)
filename = replace(filename, "/", "\")

Dim nodito As staticNode
Dim formulario As frmDoc
 
 
' si esto es un include y ya existen nodos de este archivo mandamos a armar desde el buffer
If isInclude Then
    If existTreeForFile(filename) Then
            ' vamos a ver si existe el formulario de este include
            ' y si hace falta refrescarlo para no mandarlo al lomo
            Set formulario = FindFileForm(filename)
            
            Dim refrescar As Boolean
            refrescar = False
            If Not formulario Is Nothing Then
                If formulario.mustRefresh = True Then
                    refrescar = True
                End If
            End If
            
            If refrescar = False Then
                makeProgramTree filename, True
                Exit Sub
            End If
    End If
Else
    
    ' en cambio, si no es include
    ' quiere decir que lo mandaron a refrescarse o a armar el arbol
    ' pero... vamos a ver si ya existe en buffer para no tardar al pedo
    
    Set formulario = FindFileForm(filename)
    
    If Not formulario Is Nothing Then
        
        Dim imMainPrg As Boolean
        imMainPrg = False
        
        ' vamos a ver si se mando a refrescar el PRG principal
        ' para analizar si es necesario refrescar alguno de los includes
        If Not openedProject Is Nothing Then
            If openedProject.mainSource <> "" Then
                Dim mainPath As String
                mainPath = makePathForProject(openedProject.mainSource)
                mainPath = LCase(mainPath)
                mainPath = replace(mainPath, "/", "\")
                If filename = mainPath Then
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

        ' esto es injusto porque podria haber un include q necesita must refresh
        If formulario.mustRefresh = False Then
            If existTreeForFile(filename) Then
                makeProgramTree filename
                Exit Sub
            End If
        End If
    End If
    
    frmMain.StatusBar.PanelText("MAIN") = "Collecting info about the project"
End If

Dim srcFile As New cReadFile 'clase que lee el archivo
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

' si se trata del archivo principal setea a cero
If Not isInclude Then
    inComment = False
    MainDir = FSO.GetParentFolderName(filename) & "\"
    ' vuelve los tipos de datos del usuario a cero
    ReDim typeList(0) As String
    ' limpia los macros
    Set macros = New Collection
    Set macrosNames = New Collection
End If

For Each nodito In includesNodes
    If LCase(nodito.filename) = LCase(filename) Then
        includesNodes.Remove nodito.Key
    End If
Next nodito

srcFile.filename = filename



'Recorre todas las lineas del prg
While srcFile.canRead
    'toma uma linea
    linea = srcFile.getLine
    
    lineNum = lineNum + 1
    
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
    
    ' TODO analizamos la linea y reemplazamos todos los macros que encontremos
'    Dim macro_value As String
'    For Each macro_value In macros
'    Do
'    while instr(linea,
'    Next
    
    num = 0
    num2 = 0
    
   
    If inComment = False Then
        num = InStr(linea, "//")
        
        If num > 0 Then
            linea = Mid(linea, 1, num - 1)
        End If
    End If
    
    If inComment = True Then
        num = InStr(linea, "/*")
        num2 = InStr(linea, "*/")
        
        If ((num > 0 And num2 < num) Or num = 0) And num2 > 0 Then
            linea = Mid(linea, num2 + 2)
            inComment = False
        End If
    End If
    
    If inComment = False Then
        num = InStr(linea, "/*")
        While (num > 0)
            num2 = InStr(linea, "*/")
            If num2 > num Then
                inComment = False
                linea = Mid(linea, 1, num - 1) & Mid(linea, num2 + 2)
            Else
                If inComment = False Then
                    linea = Mid(linea, 1, num - 1)
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
    '*********** ANALISIS DE CODIGO ************
    '*******************************************
    
    While Len(linea) > 0
        On Error GoTo proximaVuelta
        palabra = getWord(linea)
        
        ' vemos si la palabra no es un macro
        Dim macro_value As Variant
        Dim macindex As Integer
        Dim macroname As Variant
        
        While macindex < macros.count
            macro_value = macros.item(macindex + 1)
            macroname = macrosNames.item(macindex + 1)
            If LCase(palabra) = Trim(LCase(macroname)) Then
                palabra = ""
                linea = macro_value & " " & linea
            End If
            macindex = macindex + 1
        Wend
        
        If palabra <> "" Then
            If endsCount > 0 Then
                Select Case LCase(palabra)
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
                        ' copia de linea para tomar la siguiente palabra
                        lineaTemp = linea
                        ' ese debe ser el nombre
                        nextWord = getWord(lineaTemp)
                        ' la siguiente palabra nos interesa, tiene que ser (
                        nextWord = getWord(lineaTemp)
                        
                        If nextWord = "(" Then
                            ' si se estan por definir parametros entonces
                            ' es mas que obvio que no es una variable (es una funcion)
                            declarationType = "functionName"
                        End If
                    End If
                End If
                
                Dim Inicio As String
                Dim Largo As String
                
                        
                Select Case LCase(palabra)
                    Case "#define":
                        
                        Dim macroValue As String
                        Dim Macro As String
                        
                        Largo = InStr(linea, " ")
                        If Largo > 0 Then
                            ' puede tener valor despues del espacio
                            Macro = Trim(Left(linea, Largo))
                            
                            
                            macroValue = Trim(Mid(linea, Largo))
                        Else
                            ' no hay valor, solo nombre
                            Macro = Trim(linea)
                        End If
                        
                        If macroValue <> "" Then
                        On Error GoTo macroerror
                            macros.Add macroValue, Macro
macroerror:                 macrosNames.Add Macro, Macro
                        End If
                        
                        
                        'On Error GoTo Esquiva1

                        'ReDim Preserve typeList(UBound(typeList) + 1) As String
                        'typeList(UBound(typeList)) = LCase(Macro)

                        linea = ""
                        declarationType = ""
                        
                    Case "include":
                        
                        Dim incluir As String
                        Inicio = InStr(linea, Chr(34)) + 1
                        Largo = InStrRev(linea, Chr(34)) - Inicio
                        incluir = Mid(linea, Inicio, Largo)
                        
                        If InStr(linea, ";") <> 0 Then
                            linea = Mid(linea, InStr(linea, ";") + 1)
                        End If
                        
                        'aca sabemos si es un dir relativo
                        If FSO.GetDriveName(incluir) = "" Then
                           'entonces lo adicionamos al dir del proyecto principal
                           incluir = MainDir & incluir
                        End If
                        
                        'bueno, a ver si existe
                        If FSO.FileExists(incluir) Then
                        
                           incluir = LCase(incluir)
                           
                           
                           'On Error GoTo Esquiva1
                           
                           newKey = "INCLUDE" & "|" & incluir & "|" & filename
                                                                                 
                           Set tempNode = New staticNode
                           includesNodes.Add tempNode, newKey
                        
                           tempNode.filename = filename
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
                    Case "function": declarationType = "functionName"
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
                            varType = Left(fatherNode, InStr(fatherNode, "|") - 1)
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
                            waitFor = Chr(34) ' lo manda a no declarar hasta que se encuentre esa palabra
                        End If
                    Case "=":
                        If declarationType = "variables" And waitFor = "" Then
                            waitFor = ";" ' lo manda a no declarar hasta que se encuentre esa palabra
                        End If
                    Case "[":
                        If declarationType = "variables" And waitFor = "" Then
                            waitFor = "]" ' lo manda a no declarar hasta que se encuentre esa palabra
                        End If
                    
                    Case Else:
                        If waitFor <> "" Then ' si se espera una palabra
                            If palabra = waitFor Then ' si encuentra lo que se esperaba
                               waitFor = "" ' deja continuar con la declaración normal
                            End If
                        Else
                            Select Case declarationType
                           ' Case "macroDefinition":
                                ' arreglado en la revision del 2006
                                ' un macro NO es un tipo de dato
                                
                               ' ReDim Preserve typeList(UBound(typeList) + 1) As String
                               ' typeList(UBound(typeList)) = LCase(palabra)
                                
                               ' declarationType = ""
                                
                            Case "type":
                                
                                newKey = "type" & "|" & palabra & "|" & filename
                                
                                'On Error GoTo Esquiva2
                                
                                
                                ' agrega la declaracion al buffer general
                                
                                Set tempNode = New staticNode
                                includesNodes.Add tempNode, newKey
                                
                                tempNode.father = ""
                                tempNode.filename = filename
                                tempNode.Icon = 2
                                tempNode.Key = newKey
                                tempNode.name = palabra
                                tempNode.varType = "type"
                                
                                ' agrega el tipo de datos a la lista general
                                ReDim Preserve typeList(UBound(typeList) + 1) As String
                                typeList(UBound(typeList)) = LCase(palabra)
                                
    
                                declarationType = "variables"
    
                                fatherNode = newKey
                                fatherType = "type"
                                
Esquiva2:
'On Error GoTo Termina
                            Case "struct":
                                
                                newKey = varType & "|" & palabra & "|" & fatherNode & "|" & filename
                                
                                'On Error GoTo Esquiva3
                                
                                Set tempNode = New staticNode
                                includesNodes.Add tempNode, newKey
                                
                                structFatherType = fatherType
                                structFather = fatherNode
                                
                                
                                tempNode.father = fatherNode
                                tempNode.filename = filename
                                tempNode.Icon = 2
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
                                
                                ' si es una funcion q retorna un tipo (declarada con int nombre(), por ejemplo)
                                If isUserDefinedType(palabra) Or isDefinedType(palabra) Then
                                    ' copia de linea para tomar la siguiente palabra
                                    lineaTemp = linea
                                    ' ese debe ser el nombre
                                    nextWord = getWord(lineaTemp)
                                    newKey = "function" & "|" & nextWord & "|" & filename
                                Else
                                    'si es una declaracion comun de funcion (Function juana())
                                    newKey = "function" & "|" & palabra & "|" & filename
                                End If
                                
                               ' On Error GoTo Esquiva4
                                'TEMPORAL BUGFIX:
                                'Cuando se usan declares la función aparece repetida
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
                                    tempNode.filename = filename
                                    tempNode.Icon = 5
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
                                End If
                                
Esquiva4:
'On Error GoTo Termina
                            Case "importName":
                                
                                newKey = "import" & "|" & palabra & "|" & filename
                                
                                'On Error GoTo Esquiva5
                                
                                Set tempNode = New staticNode
                                includesNodes.Add tempNode, newKey
                                
                                tempNode.father = ""
                                tempNode.filename = filename
                                tempNode.Icon = 7
                                tempNode.Key = newKey
                                tempNode.name = palabra
                                tempNode.varType = "import"
                                
                                
                                declarationType = ""
                                
Esquiva5:
'On Error GoTo Termina
                            
                            Case "processName":
                                
                                newKey = "process" & "|" & palabra & "|" & filename
                                
                                'On Error GoTo Esquiva6
                                'TEMPORAL BUGFIX:
                                'Cuando se usan declares el proceso aparece repetido
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
                                    tempNode.filename = filename
                                    tempNode.Icon = 0
                                    tempNode.Key = newKey
                                    tempNode.name = palabra
                                    tempNode.varType = "process"
                                    tempNode.parameters = getParameters(linea)
            
                                    declarationType = ""
            
                                    fatherNode = newKey
                                    fatherType = "process"
                                End If
                                
Esquiva6:
'On Error GoTo Termina
                                
                            Case "variables":
                                ' bueno, si la palabra no es nada de parte del lenguaje...
                                If isUserDefinedType(palabra) = False And isDefinedType(palabra) = False And isOperator(palabra) = False Then
                                    
                                    Select Case varType
                                        Case "private":  imagen = 1
                                        Case "local": imagen = 3
                                        Case "global": imagen = 4
                                        Case "type": imagen = 6
                                        Case "struct": imagen = 6
                                        Case "const": imagen = 6
                                        Case Else: imagen = -1
                                    End Select
                                    
                                    newKey = lineNum & "|" & imagen & "|" & palabra & "|" & fatherNode & "|" & filename
                                    
                                    'On Error GoTo Esquiva7
                                    
                                    Set tempNode = New staticNode
                                    includesNodes.Add tempNode, newKey
                                
                                    tempNode.father = fatherNode
                                    tempNode.filename = filename
                                    tempNode.Icon = imagen
                                    tempNode.Key = newKey
                                    tempNode.name = palabra
                                    tempNode.varType = "var"
                                    tempNode.varAmbient = varType
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

Termina:
If Not isInclude Then
    frmMain.StatusBar.PanelText("MAIN") = ""
    makeProgramTree filename
End If

End Sub

' **************************************************************
'
'          TOMA UNA SOLA PALABRA DE LA LINEA DEL PROYECTO
'
' **************************************************************


Public Function getWord(ByRef linea As String) As String
    Dim num() As Variant
    Dim num2 As Variant
    Dim num3 As Variant
    Dim operatorsList() As String
    Dim returnValue As String
    
    'lista de operadores usados en la declaración
    ReDim operatorsList(LBound(allOperatorsList) To (UBound(allOperatorsList) + 1))
    
    Dim i As Long
    
    For i = LBound(allOperatorsList) To UBound(allOperatorsList)
      operatorsList(i) = allOperatorsList(i)
    Next
    operatorsList(UBound(operatorsList)) = " "
    
    ReDim num(LBound(operatorsList) To UBound(operatorsList))
    
    getWord = ""
    
    
    For i = LBound(operatorsList) To UBound(operatorsList)
        num(i) = InStr(linea, operatorsList(i))
    Next i
    
    num3 = 0
    num2 = Len(linea)
    For i = LBound(num) To UBound(num)
        If num(i) > 0 And num(i) <= num2 Then
            num2 = num(i)
            num3 = num(i)
        End If
    Next i
            
    'buscar el num mas chico
    
    If num3 > 0 Then
        'si la longitud es 1, entonces es un operador
        If num3 = 1 Then
            returnValue = Trim(Mid(linea, 1, num3))
            linea = Trim(Mid(linea, num3 + 1))
        Else
        'si no es una palabra, y le eliminamos el operador
            returnValue = Trim(Mid(linea, 1, num3 - 1))
            linea = Trim(Mid(linea, num3))
        End If
    Else
    'si no hay mas terminación de palabra es que termino la linea, asi que devolvemos la palabra como la linea
        returnValue = Trim(linea)
        linea = ""
    End If
    
    If InStr(returnValue, Chr(9)) = 1 Then
        returnValue = Mid(returnValue, 2)
    Else
        If InStr(returnValue, Chr(9)) = Len(returnValue) Then
            If Len(returnValue) = 0 Then
                returnValue = ""
            Else
                returnValue = Mid(returnValue, 1, Len(returnValue) - 1)
            End If
        End If
    End If
    getWord = returnValue
End Function

' **************************************************************
'
'       TOMA UNA SOLA PALABRA DE LA LINEA DEL PROYECTO
'                Desde atras para adelante
'
' **************************************************************


Public Function getWordRev(ByRef linea As String) As String
    Dim num() As Long
    Dim num2 As Variant
    Dim num3 As Variant
    Dim operatorsList() As String
    Dim returnValue As String
    Dim i As Integer
    
    ' elimina caracteres indeseados
    linea = Trim(replace(linea, Chr(9), " "))
    
    'lista de operadores usados en la declaración
    ReDim operatorsList(LBound(allOperatorsList) To (UBound(allOperatorsList) + 1))
    
    For i = LBound(allOperatorsList) To UBound(allOperatorsList)
      operatorsList(i) = allOperatorsList(i)
    Next
    operatorsList(UBound(operatorsList)) = " "
    
    ReDim num(LBound(operatorsList) To UBound(operatorsList))
    
    getWordRev = ""
    
    i = 0
    
    For i = LBound(operatorsList) To UBound(operatorsList)
        num(i) = InStrRev(linea, operatorsList(i))
    Next i
    
    'buscar el num mas grande
    num3 = 0
    num2 = 0
    For i = LBound(num) To UBound(num)
        If num(i) > num2 Then
            num2 = num(i)
            num3 = num(i)
        End If
    Next i
    
    
    If num3 > 0 Then
        'si la la ocorrencia es igual a la longitud de la linea, entonces es un operador
        
        If num3 = Len(linea) Then
            returnValue = Right(linea, 1)
            linea = Trim(Mid(linea, 1, num3 - 1))
        Else
            returnValue = Trim(Mid(linea, num3 + 1))
            linea = Trim(Mid(linea, 1, num3))
        End If
    Else
    'si no hay mas terminación de palabra es que termino la linea, asi que devolvemos la palabra como la linea
        returnValue = Trim(linea)
        linea = ""
    End If
        
    getWordRev = returnValue
End Function

Public Function isUserDefinedType(ByVal palabra As String) As Boolean
    isUserDefinedType = False
    'si no es un tipo definido por el usuario
    If Not IsEmpty(LBound(typeList)) Then
        Dim i As Long
        For i = LBound(typeList) To UBound(typeList)
            If LCase(palabra) = LCase(typeList(i)) Then
                isUserDefinedType = True
                Exit Function
            End If
        Next i
    End If
End Function

Public Function isDefinedType(ByVal palabra As String) As Boolean
    
Dim num As Double

isDefinedType = False

For num = LBound(allTypeList) To UBound(allTypeList)
    If LCase(palabra) = LCase(allTypeList(num)) Then
        isDefinedType = True
        Exit Function
    End If
Next num

End Function

' se fija en el fdl.lan por los operadores y si la palabra es uno de ellos

Public Function isOperator(ByVal palabra As String) As Boolean
Dim num As Double

isOperator = False

For num = LBound(allOperatorsList) To UBound(allOperatorsList)
    If LCase(palabra) = LCase(allOperatorsList(num)) Then
        isOperator = True
        Exit Function
    End If
Next num
End Function



' devuelve true si la palabra pasada como parametro forma parte del
' lenguaje

Public Function isReservedWord(ByVal palabra As String) As Boolean
Dim linea As String
Dim num As Double

num = FreeFile()

isReservedWord = False
If isUserDefinedType(palabra) Then
    isReservedWord = True
    Exit Function
End If
Open App.Path & "\Help\fdl.lan" For Input As #num
    Do Until EOF(num)
        'lee una linea
        Line Input #num, linea
        ' evita comentarios
        If InStr(linea, "//#") <> 1 And linea <> "" Then
            If LCase(palabra) = LCase(Trim(linea)) Then
                isReservedWord = True
                Close #num
                Exit Function
            End If
        End If
    Loop
Close #num

End Function
