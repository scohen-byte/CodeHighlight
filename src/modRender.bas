Attribute VB_Name = "modRender"
'==============================================================================
' modRender - applies the scanner's spans to a shape.
'
' The whole module is about one loop, and two things make it correct:
'
'   RESET FIRST, THEN COLOUR. Setting the entire range to the default colour
'   before applying spans is what makes re-highlighting idempotent. Without it,
'   editing a line and pressing Stylize again leaves the old colours behind
'   on the characters that no longer belong to that token.
'
'   TOKENIZE THE STRING YOU ARE ABOUT TO INDEX INTO. Positions come back from
'   the scanner 1-based, and TextRange.Characters is 1-based, but only because
'   both refer to the SAME string - the one PowerPoint just handed back. Reading
'   the text from anywhere else (a file, a cached copy with LF endings) shifts
'   every colour by one character per line, which looks bizarre and is the
'   classic bug here. PLAN.md section 13.
'==============================================================================
Option Explicit

' Returns the number of colour assignments actually made, which is the number of
' COM round trips and therefore the thing worth measuring.
Public Function ApplyHighlight(ByVal shp As Shape, ByVal langId As String) As Long
    Dim tr As TextRange
    Dim spans() As Span, n As Long, i As Long, applied As Long
    Dim lineOf() As Long, emph As String, dimming As Boolean
    Dim c As Long

    Set tr = shp.TextFrame.TextRange
    If Len(tr.text) = 0 Then Exit Function

    tr.Font.Name = THEME_FONT
    tr.Font.Color.RGB = ThemeColor(tkDefault)

    ' Emphasis BEFORE the colours. Setting Font.Highlight across the whole range
    ' collapses the runs, so doing it afterwards repaints every character in one
    ' colour and throws the syntax highlighting away.
    ApplyEmphasis shp

    ' When some lines are emphasised, every OTHER line is dimmed. The band alone
    ' is easy to miss from the back of a lecture hall; fading the surroundings
    ' makes the live lines carry the slide.
    emph = modBlock.GetEmphasis(shp)
    dimming = (Len(emph) > 0)
    If dimming Then
        BuildLineIndex tr.text, lineOf
        tr.Font.Color.RGB = ThemeDimmed(ThemeColor(tkDefault))
    End If

    n = modLexer.Tokenize(tr.text, modLangRegistry.GetLang(langId), spans)

    For i = 0 To n - 1
        ' Default-coloured spans were already handled by the reset above, and
        ' skipping them removes roughly half the round trips.
        c = ThemeColor(spans(i).Kind)
        If dimming Then
            If Not LineIsEmphasised(lineOf, spans(i).Start, emph) Then c = ThemeDimmed(c)
        End If

        ' The reset already painted the default colour, so a default span only
        ' needs touching when dimming has changed what "default" should be.
        If spans(i).Kind <> tkDefault Or dimming Then
            If c <> tr.Characters(spans(i).Start, 1).Font.Color.RGB Then
                tr.Characters(spans(i).Start, spans(i).Length).Font.Color.RGB = c
                applied = applied + 1
            End If
        End If
    Next i

    ApplyHighlight = applied
End Function

' char index -> line number, built once per render. Walking the text per span
' would be quadratic, and a long block has hundreds of spans.
Private Sub BuildLineIndex(ByVal text As String, ByRef lineOf() As Long)
    Dim i As Long, n As Long, ch As String

    ReDim lineOf(1 To IIf(Len(text) < 1, 1, Len(text)))
    n = 1
    For i = 1 To Len(text)
        lineOf(i) = n
        ch = Mid$(text, i, 1)
        If ch = vbCr Or ch = vbLf Or ch = Chr$(11) Then n = n + 1
    Next i
End Sub

Private Function LineIsEmphasised(ByRef lineOf() As Long, ByVal charIdx As Long, _
                                  ByVal emph As String) As Boolean
    Dim ln As Long
    On Error Resume Next
    If charIdx < LBound(lineOf) Or charIdx > UBound(lineOf) Then Exit Function
    ln = lineOf(charIdx)
    LineIsEmphasised = (InStr("," & emph & ",", "," & CStr(ln) & ",") > 0)
End Function

' Paints the band behind whichever lines the block says are emphasised, and
' clears it everywhere else.
'
' Called from ApplyHighlight so emphasis survives every re-render. That matters
' more than it sounds: the whole use case is duplicating a slide and moving the
' emphasis down a line, which means Stylize runs constantly and must not undo
' it.
Public Sub ApplyEmphasis(ByVal shp As Shape)
    Dim spec As String, parts() As String, i As Long
    Dim txt As String, startIdx As Long, length As Long

    On Error GoTo Done

    ' Clear by painting the band the same colour as the block.
    '
    ' Font.Highlight is a ColorFormat - it has RGB, and NO Visible, so there is
    ' no "off". On a solid dark block, a band in the background colour is
    ' indistinguishable from no band at all.
    shp.TextFrame2.TextRange.Font.Highlight.RGB = ThemeBackColor()

    spec = modBlock.GetEmphasis(shp)
    If Len(spec) = 0 Then Exit Sub

    txt = shp.TextFrame.TextRange.text
    parts = Split(spec, ",")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            modBlock.LineCharRange txt, CLng(Val(parts(i))), startIdx, length
            ' A zero length is a blank line - still worth banding, so the
            ' emphasis reads as a contiguous block rather than a dashed one.
            If startIdx > 0 Then
                If length = 0 Then length = 1
                shp.TextFrame2.TextRange.Characters(startIdx, length) _
                   .Font.Highlight.RGB = ThemeEmphasisColor()
            End If
        End If
    Next i
Done:
End Sub
