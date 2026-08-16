Attribute VB_Name = "modSwatch"
'==============================================================================
' modSwatch - solid colour squares for the ribbon, made at runtime.
'
' The ribbon can show an image beside a gallery item, but only as an
' IPictureDisp. There is no way to say "a square of this colour" in the XML, and
' imageMso ids cannot be tinted. So the picture has to come from somewhere, and
' the only offline route from a colour to an IPictureDisp is a file that
' LoadPicture can open.
'
' A 24-bit BMP is the least machinery that works: fifty-four bytes of header and
' then three bytes per pixel, no compression, no palette, no library. PNG would
' need a deflate stream and a CRC.
'
' The files go to the user's TEMP folder and are cached for the session. A
' ribbon asks for every item image again on each Invalidate, so regenerating
' them per call would write to disk every time the selection changes.
'
' BMP is bottom-up and rows are padded to a multiple of four bytes. Neither
' matters for a solid square of a width divisible by four, but assuming that
' silently is how a swatch ends up skewed the first time the size changes.
'==============================================================================
Option Explicit

Private Const PX As Long = 16                   ' the ribbon draws these at 16x16

Private mCache      As Object                   ' colour -> IPictureDisp
Private mCacheReady As Boolean

' Nothing when the swatch cannot be made, which the caller treats as "no image"
' rather than as an error. A gallery with labels and no pictures still works.
Public Function Swatch(ByVal rgbColor As Long) As Object
    Dim key As String, path As String

    On Error GoTo Failed

    If Not mCacheReady Then
        Set mCache = CreateObject("Scripting.Dictionary")
        mCacheReady = True
    End If

    key = CStr(rgbColor)
    If mCache.Exists(key) Then
        Set Swatch = mCache(key)
        Exit Function
    End If

    path = SwatchPath(rgbColor)
    If Len(Dir$(path)) = 0 Then WriteBmp path, rgbColor
    Set Swatch = LoadPicture(path)
    mCache.Add key, Swatch
    Exit Function

Failed:
    Set Swatch = Nothing
End Function

Private Function SwatchPath(ByVal rgbColor As Long) As String
    Dim t As String
    t = Environ$("TEMP")
    If Len(t) = 0 Then t = Environ$("TMP")
    If Right$(t, 1) <> "\" Then t = t & "\"
    SwatchPath = t & "CodeHighlight-swatch-" & CStr(rgbColor) & "-" & CStr(PX) & ".bmp"
End Function

' VBA's RGB() packs BLUE in the high byte, which is the opposite of what the
' name suggests and the opposite of what a BMP pixel wants. Unpacking here, once,
' keeps that fact in one place.
Private Sub WriteBmp(ByVal path As String, ByVal rgbColor As Long)
    Dim f As Integer, i As Long
    Dim rowBytes As Long, padding As Long, pixelBytes As Long, fileSize As Long
    Dim r As Byte, g As Byte, b As Byte, row() As Byte

    r = CByte(rgbColor And &HFF&)
    g = CByte((rgbColor \ &H100&) And &HFF&)
    b = CByte((rgbColor \ &H10000) And &HFF&)

    rowBytes = PX * 3
    padding = (4 - (rowBytes Mod 4)) Mod 4
    pixelBytes = (rowBytes + padding) * PX
    fileSize = 54 + pixelBytes

    ReDim row(0 To rowBytes + padding - 1)
    For i = 0 To PX - 1
        ' BGR, not RGB. A BMP pixel is stored blue first.
        row(i * 3) = b
        row(i * 3 + 1) = g
        row(i * 3 + 2) = r
    Next i

    f = FreeFile
    Open path For Binary Access Write As #f

    ' BITMAPFILEHEADER, 14 bytes.
    Put #f, , CByte(66): Put #f, , CByte(77)        ' "BM"
    Put #f, , fileSize                              ' Long, 4 bytes
    Put #f, , CLng(0)                               ' reserved
    Put #f, , CLng(54)                              ' offset to the pixels

    ' BITMAPINFOHEADER, 40 bytes.
    Put #f, , CLng(40)                              ' header size
    Put #f, , CLng(PX)                              ' width
    Put #f, , CLng(PX)                              ' height, positive = bottom-up
    Put #f, , CInt(1)                               ' planes
    Put #f, , CInt(24)                              ' bits per pixel
    Put #f, , CLng(0)                               ' BI_RGB, no compression
    Put #f, , pixelBytes
    Put #f, , CLng(2835)                            ' 72 dpi, in pixels per metre
    Put #f, , CLng(2835)
    Put #f, , CLng(0)                               ' colours used
    Put #f, , CLng(0)                               ' colours important

    For i = 1 To PX
        Put #f, , row
    Next i

    Close #f
End Sub
