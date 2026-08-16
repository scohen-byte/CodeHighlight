Attribute VB_Name = "modSelfTest"
'==============================================================================
' modSelfTest - runs the scanner over the sample corpus and dumps what it
'               decided, so it can be diffed against tools/lexref.py.
'
' This module exists to make the lexer testable WITHOUT PowerPoint in the loop
' as a human step. tools/run-lexer-tests.sh drives it over COM: it imports the
' modules into a scratch deck, calls MaskFileToFile once per sample, and diffs
' the result against the reference. A lexer bug shows up as a wrong column in a
' diff rather than as a vague sense that the colours look off.
'==============================================================================
Option Explicit

' The language for this run. Set once with SetLanguage rather than passed on
' every call, because Application.Run proved unreliable with three arguments
' from PowerShell - two always work, three sometimes report "Sub or function
' not defined" even though the procedure is plainly there. Not root-caused;
' the dependency is simply removed.
Private mLangId As String

Public Sub SetLanguage(ByVal langId As String)
    mLangId = langId
End Sub

Private Function RunLangId() As String
    If Len(mLangId) = 0 Then mLangId = modLangRegistry.DefaultLangId()
    RunLangId = mLangId
End Function

Private Function RunLang() As LangDef
    If Len(mLangId) = 0 Then mLangId = modLangRegistry.DefaultLangId()
    RunLang = modLangRegistry.GetLang(mLangId)
End Function

' Smoke test for the harness itself. VBA compiles on demand, so a module that
' fails to compile reports as "Sub or function not defined" when you call into
' it - which reads exactly like a wrong macro name. Calling this first separates
' the two: if Ping works and the rest do not, the module is broken, not the
' plumbing.
Public Function Ping() As String
    Ping = "pong " & modLangRegistry.LangCount() & " lang(s)"
End Function

' Masks EVERY sample in a folder in a single call, and returns a summary.
'
' One macro call for the whole corpus rather than three per sample. Application.Run
' from PowerShell proved unreliable when called repeatedly against a freshly
' imported project - procedures that were demonstrably present reported as
' "Sub or function not defined". Doing the whole run inside VBA sidesteps it, and
' is faster anyway: the COM boundary is crossed once instead of thirty times.
'
' Returns one line per sample: "<name> <maskLen> <spanCount>", or
' "<name> ERROR <description>".
Public Function RunCorpus(ByVal sampleDir As String, ByVal outDir As String) As String
    Dim names() As String, count As Long, i As Long
    Dim fileName As String, src As String, mask As String, report As String
    Dim spans() As Span, n As Long
    Dim t0 As Single, totalChars As Long

    ReDim names(0 To 255)
    count = 0

    ' Collect the names first. Interleaving Dir() with the file reads below
    ' would clobber its iteration state.
    fileName = Dir(sampleDir & "\*.py")
    Do While Len(fileName) > 0
        If count > UBound(names) Then ReDim Preserve names(0 To count + 255)
        names(count) = fileName
        count = count + 1
        fileName = Dir()
    Loop

    t0 = Timer
    For i = 0 To count - 1
        On Error GoTo ItemFailed
        src = ReadTextFile(sampleDir & "\" & names(i))
        mask = modLexer.MaskOf(src, RunLang())
        n = modLexer.Tokenize(src, RunLang(), spans)
        WriteTextFile outDir & "\" & BaseName(names(i)) & ".mask", mask
        report = report & names(i) & " " & Len(mask) & " " & n & vbLf
        totalChars = totalChars + Len(src)
        GoTo NextItem
ItemFailed:
        report = report & names(i) & " ERROR " & Err.Description & vbLf
        Err.Clear
NextItem:
        On Error GoTo 0
    Next i

    ' Timing rides along here rather than in its own macro call. Repeated
    ' Application.Run calls against a freshly imported project proved unreliable,
    ' so the whole run is kept to as few crossings as possible.
    report = report & "TOTAL " & CLng((Timer - t0) * 1000) & " " & totalChars & vbLf

    RunCorpus = report
End Function

Private Function BaseName(ByVal fileName As String) As String
    Dim dot As Long
    dot = InStrRev(fileName, ".")
    If dot > 0 Then BaseName = Left$(fileName, dot - 1) Else BaseName = fileName
End Function

' THE THIN SLICE: insert one block on a fresh slide, highlight it, export a PNG.
'
' Everything in one macro call, both because repeated Application.Run calls
' proved unreliable (section 4 of HANDOFF.md) and because the timings are only
' meaningful measured inside VBA.
'
' Also writes the mask of the SHAPE'S OWN TEXT next to the PNG. That is the
' check that matters here: if PowerPoint altered the text on the way in - line
' endings, tabs, autocorrected quotes - the shape mask and the file mask
' disagree, and every colour is shifted. Comparing renders by eye would not
' catch a one-character shift, but a diff does.
Public Function BuildSlice(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim code As String, shapeText As String, report As String
    Dim t0 As Single, buildMs As Long, colourMs As Long
    Dim applied As Long, dot As Long

    On Error GoTo Failed

    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)

    t0 = Timer
    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, RunLangId())
    buildMs = CLng((Timer - t0) * 1000)

    t0 = Timer
    applied = modRender.ApplyHighlight(shp, RunLangId())
    colourMs = CLng((Timer - t0) * 1000)

    ' The mask of what is actually in the shape, not of what was in the file.
    shapeText = shp.TextFrame.TextRange.text
    dot = InStrRev(pngPath, ".")
    If dot > 0 Then
        WriteTextFile Left$(pngPath, dot - 1) & ".mask", _
                      modLexer.MaskOf(shapeText, RunLang())
    End If

    sld.Export pngPath, "PNG", 1920, 1080

    BuildSlice = "lines=" & modBlock.CountLines(shapeText) & _
                 " shapechars=" & Len(shapeText) & _
                 " filechars=" & Len(code) & _
                 " applied=" & applied & _
                 " build_ms=" & buildMs & _
                 " colour_ms=" & colourMs & _
                 " height=" & Format$(shp.Height, "0.0") & _
                 " top=" & Format$(shp.Top, "0.0")
    Exit Function

Failed:
    BuildSlice = "ERROR " & Err.Number & ": " & Err.Description
End Function

' One file, for poking at a single sample from the VBA editor's Immediate
' window. The test harness uses RunCorpus instead - see the note there.
' Returns the mask length, or -1 on failure.
Public Function MaskFileToFile(ByVal inPath As String, ByVal outPath As String) As Long
    Dim src As String, mask As String

    On Error GoTo Failed
    src = ReadTextFile(inPath)
    mask = modLexer.MaskOf(src, RunLang())
    WriteTextFile outPath, mask
    MaskFileToFile = Len(mask)
    Exit Function

Failed:
    WriteTextFile outPath, "ERROR " & Err.Number & ": " & Err.Description
    MaskFileToFile = -1
End Function

'------------------------------------------------------------------------------

' Line endings are normalised to LF, matching Python's universal newlines, so
' character positions agree with what the reference tool saw. Without this every
' mask would be one character per line out of step.
Private Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer, buf As String
    f = FreeFile
    Open path For Binary Access Read As #f
    buf = Space$(LOF(f))
    Get #f, , buf
    Close #f
    buf = Replace(buf, vbCrLf, vbLf)
    buf = Replace(buf, vbCr, vbLf)
    ReadTextFile = buf
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, content;
    Close #f
End Sub
