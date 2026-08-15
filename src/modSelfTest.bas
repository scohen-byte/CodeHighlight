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

' Returns the mask length, or -1 on failure. Writing the error into the output
' file too, so a failure is visible in the diff instead of looking like an
' empty result.
Public Function MaskFileToFile(ByVal inPath As String, ByVal outPath As String, _
                               ByVal langId As String) As Long
    Dim src As String, mask As String

    On Error GoTo Failed
    src = ReadTextFile(inPath)
    mask = modLexer.MaskOf(src, modLangRegistry.GetLang(langId))
    WriteTextFile outPath, mask
    MaskFileToFile = Len(mask)
    Exit Function

Failed:
    WriteTextFile outPath, "ERROR " & Err.Number & ": " & Err.Description
    MaskFileToFile = -1
End Function

' Milliseconds to tokenize a file, for the "is this fast enough" question in
' PLAN section 7. Returns -1 on failure.
Public Function TimeTokenize(ByVal inPath As String, ByVal langId As String, _
                             ByVal reps As Long) As Long
    Dim src As String, spans() As modLexer.Span, i As Long
    Dim t0 As Single, lang As LangDef

    On Error GoTo Failed
    src = ReadTextFile(inPath)
    lang = modLangRegistry.GetLang(langId)
    t0 = Timer
    For i = 1 To reps
        modLexer.Tokenize src, lang, spans
    Next i
    TimeTokenize = CLng((Timer - t0) * 1000)
    Exit Function

Failed:
    TimeTokenize = -1
End Function

' Span count for a file, so the renderer's COM round-trip cost can be predicted
' before the renderer exists.
Public Function SpanCountOf(ByVal inPath As String, ByVal langId As String) As Long
    Dim src As String, spans() As modLexer.Span
    On Error GoTo Failed
    src = ReadTextFile(inPath)
    SpanCountOf = modLexer.Tokenize(src, modLangRegistry.GetLang(langId), spans)
    Exit Function
Failed:
    SpanCountOf = -1
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
