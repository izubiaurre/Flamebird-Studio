Attribute VB_Name = "modZlib"
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

' zlib.h -- interface of the 'zlib' general purpose compression library
'  version 1.0.4, Jul 24th, 1996.

' If you do not have any name conflicts with compress(), uncompress() or crc32()
' then you can set the following flag
#Const ZLIB_SHORT_NAMES = 1
' If you KNOW that all your uncompressed data is ASCII strings, then
' use can set the following flag. After decompression, don't forget to shorten
' your data buffer to reflect the decompressed length
#Const ZLIB_ASCII_ONLY = 0

'            API
' compress(dest, destLen, source, sourceLen)
'   Compresses the source buffer into the destination buffer.  sourceLen is
'   the byte length of the source buffer. Upon entry, destLen is the total
'   size of the destination buffer, which must be at least 0.1% larger than
'   sourceLen plus 12 bytes. Upon exit, destLen is the actual size of the
'   compressed buffer.
' uncompress(dest, destLen, source, sourceLen)
'   Decompresses the source buffer into the destination buffer.  sourceLen is
'   the byte length of the source buffer. Upon entry, destLen is the total
'   size of the destination buffer, which must be large enough to hold the
'   entire uncompressed data. (The size of the uncompressed data must have
'   been saved previously by the compressor and transmitted to the decompressor
'   by some mechanism outside the scope of this compression library.)
'   Upon exit, destLen is the actual size of the compressed buffer.
' gzopen(path, mode)
'   Opens a gzip (.gz) file for reading or writing. The mode parameter
'   is as in fopen ("rb" or "wb") but can also include a compression level
'   ("wb9").  gzopen can be used to read a file which is not in gzip format;
'   in this case gzread will directly read from the file without decompression.
' gzread(file, buf, len)
'   Reads the given number of uncompressed bytes from the compressed file.
'   If the input file was not in gzip format, gzread copies the given number
'   of bytes into the buffer. gzread() returns the number of uncompressed
'   bytes actually read (0 for end of file, -1 for error)
' gzwrite(file, buf, len)
'   Writes the given number of uncompressed bytes into the compressed file.
'   gzwrite returns the number of uncompressed bytes actually written
'   (0 in case of error).
' gzflush(file, flush)
'   Flushes all pending output into the compressed file. The parameter
'   flush is 3 to restart the dictionaries and 4 to finish the file.
'   The return value is the zlib error number; you may call gzerror().
'   gzflush returns Z_OK if the flush parameter is 4 and all output could be flushed.
'   gzflush(file,3) provides "segmenting" within the compressed file, but
'   should be called only when strictly necessary because it can degrade compression.
' gzclose(file)
'   Flushes all pending output if necessary, closes the compressed file
'   and deallocates all the (de)compression state. The return value is the zlib
'   error number; you may call gzerror().
' gzerror(file, errnum)
'   Returns the error message for the last error which occurred on the
'   given compressed file. errnum is set to zlib error number. If an
'   error occurred in the file system and not in the compression library,
'   errnum is set to Z_ERRNO(-1)
' adler32(adler, buf, len)
'   Update a running Adler-32 checksum with the bytes buf[0..len-1] and
'   return the updated checksum. If buf is NULL, this function returns
'   the required initial value for the checksum.
'   An Adler-32 checksum is almost as reliable as a CRC32 but can be computed
'   much faster
' crc32(crc, buf, len)
'   Update a running crc with the bytes buf[0..len-1] and return the updated
'   crc. If buf is NULL, this function returns the required initial value
'   for the crc. Pre- and post-conditioning (one's complement) is performed
'   within this function so it shouldn't be done by the application.
'   usage example:
'
'     Dim sNull as String
'     Dim crc as Long
'     crc = crc32(0, sNull, 0)      'init the crc
'     crc = crc32(crc, buffer, length)
'     crc = crc32(crc, buffer2, length2}
'     If (crc != original_crc) Then error()




Global Const Z_OK = 0
Global Const Z_STREAM_END = 1
Global Const Z_NEED_DICT = 2
Global Const Z_ERRNO = (-1)
Global Const Z_STREAM_ERROR = (-2)
Global Const Z_DATA_ERROR = (-3)
Global Const Z_MEM_ERROR = (-4)
Global Const Z_BUF_ERROR = (-5)
Global Const Z_VERSION_ERROR = (-6)
' Return codes for the compression/decompression functions. Negative
' values are errors, positive values are used for special but normal events.
' compress() returns Z_OK if success, Z_MEM_ERROR if there was not
'     enough memory, Z_BUF_ERROR if there was not enough room in the output
'     buffer
' uncompress() returns Z_OK if success, Z_MEM_ERROR if there was not
'     enough memory, Z_BUF_ERROR if there was not enough room in the output
'     buffer, or Z_DATA_ERROR if the input data was corrupted
' gzopen() returns NULL if the file could not be opened or if there was
'     insufficient memory to allocate the (de)compression state; errno
'     can be checked to distinguish the two cases (if errno is zero, the
'     zlib error is Z_MEM_ERROR).
'



'   The data format used by the zlib library is described by RFCs (Request for
'  Comments) 1950 to 1952 in the files ftp://ds.internic.net/rfc/rfc1950.txt
'  (zlib format), rfc1951.txt (deflate format) and rfc1952.txt (gzip format).

'copy-pasted directly
Public Declare Function gzReadStr Lib "zlibvb.dll" _
    Alias "gzread" (ByVal n As Long, ByVal ptr As String, ByVal lenght As Long) As Long
    
#If ZLIB_SHORT_NAMES Or ZLIB_ASCII_ONLY Then
#If ZLIB_ASCII_ONLY Then
Public Declare Function compress Lib "zlibvb.dll" (ByVal dest As String, destLen As Long, _
        ByVal Source As String, ByVal sourceLen As Long) As Long
Public Declare Function uncompress Lib "zlibvb.dll" (ByVal dest As String, destLen As Long, _
        ByVal Source As String, ByVal sourceLen As Long) As Long
#Else
Public Declare Function compress Lib "zlibvb.dll" (dest As Byte, destLen As Long, _
        Source As Byte, ByVal sourceLen As Long) As Long
Public Declare Function uncompress Lib "zlibvb.dll" (dest As Byte, destLen As Long, _
        Source As Byte, ByVal sourceLen As Long) As Long
#End If 'ZLIB_ASCII_ONLY

Public Declare Function gzopen Lib "zlibvb.dll" (ByVal Path As String, ByVal Mode As String) As Long
'Public Declare Function ZlibGzopen Lib "zlibvb.dll" Alias "gzopen" (ByVal path As String, _
        ByVal mode As String) As Long
        
#If ZLIB_ASCII_ONLY Then
Public Declare Function gzread Lib "zlibvb.dll" (ByVal file As Long, _
        ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function gzwrite Lib "zlibvb.dll" (ByVal file As Long, _
        ByVal buf As String, ByVal buf_len As Long) As Long
Public Declare Function gzReadStr Lib "zlibvb.dll" _
    Alias "gzread" (ByVal n As Long, ByVal ptr As String, ByVal lenght As Long) As Long
#Else
Public Declare Function gzread Lib "zlibvb.dll" (ByVal file As Long, _
        buf As Any, ByVal buf_len As Long) As Long
Public Declare Function gzwrite Lib "zlibvb.dll" (ByVal file As Long, buf As Any, _
        ByVal buf_len As Long) As Long

#End If 'ZLIB_ASCII_ONLY
Public Declare Function gzseek Lib "zlibvb.dll" (ByVal file As Long, ByVal off As Long, ByVal s As Long) As Long
Public Declare Function gztell Lib "zlibvb.dll" (ByVal file As Long) As Long
Public Declare Function gzeof Lib "zlibvb.dll" (ByVal file As Long) As Long
Public Declare Function gzflush Lib "zlibvb.dll" (ByVal file As Long, _
        ByVal flush As Long) As Long
Public Declare Function gzclose Lib "zlibvb.dll" (ByVal file As Long) As Long
Public Declare Function adler32 Lib "zlibvb.dll" (ByVal adler As Long, buf As Byte, _
        ByVal buf_len As Long) As Long
Public Declare Function crc32 Lib "zlibvb.dll" (ByVal crc As Long, buf As Byte, _
        ByVal buf_len As Long) As Long


' If you have a problem with the above names conflicting with other names
' in your project, you can use these longer names
#Else
Public Declare Function ZlibCompress Lib "zlibvb.dll" Alias "compress" (dest As Byte, destLen As Long, _
        Source As Byte, ByVal sourceLen As Long) As Long
Public Declare Function ZlibUncompress Lib "zlibvb.dll" Alias "uncompress" (dest As Byte, destLen As Long, _
        Source As Byte, ByVal sourceLen As Long) As Long
Public Declare Function ZlibGzopen Lib "zlibvb.dll" Alias "gzopen" (ByVal Path As String, _
        ByVal Mode As String) As Long
Public Declare Function ZlibGzread Lib "zlibvb.dll" Alias "gzread" (ByVal file As Long, _
        buf As Byte, ByVal buf_len As Long) As Long
Public Declare Function ZlibGzwrite Lib "zlibvb.dll" Alias "gzwrite" (ByVal file As Long, buf As Byte, _
        ByVal buf_len As Long) As Long
Public Declare Function ZlibGzflush Lib "zlibvb.dll" Alias "gzflush" (ByVal file As Long, _
        ByVal flush As Long) As Long
Public Declare Function ZlibGzclose Lib "zlibvb.dll" Alias "gzclose" (ByVal file As Long) As Long
Public Declare Function ZlibAdler32 Lib "zlibvb.dll" Alias "adler32" (ByVal adler As Long, buf As Byte, _
        ByVal buf_len As Long) As Long
Public Declare Function ZlibCrc32 Lib "zlibvb.dll" Alias "crc32" (ByVal crc As Long, buf As Byte, _
        ByVal buf_len As Long) As Long
#End If
        

