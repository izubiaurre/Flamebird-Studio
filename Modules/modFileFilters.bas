Attribute VB_Name = "modFileFilters"
Option Explicit

' DEFINITIONS:
' A File Type is defined by an extension an a description
'   Example: {prg, Fenix source}
'            {png, Portable network graphics}
'            {map, Old fenix bitmap}
' A FileFilter is a group of FileTypes which have something in common
'   Example: {All known graphic formats, (map, png)



Private m_FileFilters As New Dictionary

Public Function getFilter(ParamArray keys()) As String
    Dim res As String
    Dim filters() As String
    Dim Key As Variant
    Dim i As Integer
    
    On Error GoTo errhandler
    
    If UBound(keys) >= LBound(keys) Then
        For Each Key In keys 'Create the filter for each key
            filters() = Split(m_FileFilters(Key)(1), "|")
            If UBound(filters) < LBound(filters) Then
                Err.Raise 600
            End If
            For i = LBound(filters) To UBound(filters)
                filters(i) = "*." + filters(i)
            Next
            res = m_FileFilters(Key)(0) + " (" & Join(filters, ", ") & ")" & "|" & Join(filters, ";") + "|"
        Next
        res = Left(res, Len(res) - 1) 'Remove the last | symbol
    Else
        Err.Raise 600
    End If
    
    getFilter = res
    Exit Function
    
errhandler:
    If Err.Number > 0 Then ShowError ("modFileFilters.getFilter")
End Function

Private Function composeExtensions(ParamArray keys())
    Dim Key As Variant
    Dim res As String
    
    On Error GoTo errhandler
    
    If UBound(keys) >= LBound(keys) Then
        For Each Key In keys
            res = res & "|" & m_FileFilters(Key)(1) 'create an string containing all filters
        Next
        res = Right(res, Len(res) - 1) 'remove the first "|"
    Else
        Err.Raise 600
    End If
    
    composeExtensions = res
    Exit Function
    
errhandler:
    If Err.Number > 0 Then ShowError ("modFileFilters.comPoseFileFilters")
End Function

Private Sub addFileFilter(Key As String, Description As String, Filter As String)
    Dim s(1) As String
    
    s(0) = Description
    s(1) = Filter
    
    m_FileFilters.Add Key, s
End Sub

Public Sub CreateFileFilters()
    ' Create the file filters
    addFileFilter "FBP", "Flamebird MX project", "fbp"
    addFileFilter "SOURCE", "Source files", "prg|h|inc"
    
    addFileFilter "PAL", "Fenix old palette format", "pal"
    addFileFilter "FPL", "Fenix palette", "fpl"
    addFileFilter "PALETTE", "All known palette files", "pal"
    
    addFileFilter "MAP", "Fenix bitmap", "map"
    addFileFilter "FBM", "Fenix bitmap", "fbm"
    addFileFilter "PNG", "Portable nerwork graphics", "png"
    addFileFilter "BMP", "Windows bitmap", "bmp"
    addFileFilter "JPG", "JPEG Image", "jpg"
    addFileFilter "GIF", "CompuServe GIF", "gif"
    addFileFilter "IMPORTABLE_GRAPHICS", "Importable graphic files", composeExtensions("PNG", "BMP", "JPG", "GIF")
    addFileFilter "GRAPHIC_FILES", "All graphic files", composeExtensions("MAP", "FBM", "PNG", "BMP", "JPG", "GIF")
    
    addFileFilter "FPG", "Fenix graphic collection", "fpg"
    addFileFilter "FGC", "Fenix graphic collection", "fgc"
    addFileFilter "GRAPHIC_COLLECTIONS", "All graphic collections", composeExtensions("FPG", "FGC")
    
    addFileFilter "FNT", "Fenix font file", "fnt"

    addFileFilter "MOD", "Mod", "mod"
    addFileFilter "S3M", "S3m", "s3m"
    addFileFilter "XM", "Xm", "xm"
    addFileFilter "IT", "Impulse Tracker file", "it"
    addFileFilter "MID", "Midi", "mid"
    addFileFilter "MODULES", "All known song modules", composeExtensions("MOD", "S3M", "XM", "IT", "MID")
    addFileFilter "OGG", "Ogg Vorbis stream file", "ogg"
    addFileFilter "WAV", "Wave audio file", "wav"
    addFileFilter "STREAMS", "All known audio stream files", composeExtensions("OGG", "WAV")
    'addFileFilter "MODULES", "All known song modules", "mod|s3m|xm|it|mid"
    'addFileFilter "STREAM", "All known audio stream files", "ogg|mp3|wav"
    addFileFilter "SOUND_FILES", "All sound files", composeExtensions("MODULES", "STREAMS")

    addFileFilter "READABLE_FILES", "All readable files", _
                composeExtensions("FBP", "SOURCE", "PALETTE", "MAP", "FPG", "FNT", "MODULES", "STREAMS")
    
    addFileFilter "COMMON_FILES", "All common files", composeExtensions("SOURCE", "PALETTE", _
                "GRAPHIC_FILES", "GRAPHIC_COLLECTIONS", "SOUND_FILES")
End Sub
