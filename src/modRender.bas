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
    ' Reset with the rest, so turning the option off actually unbolds the line
    ' rather than leaving the last walkthrough's bold behind.
    tr.Font.Bold = msoFalse

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

    ' The transcript lines are repainted AFTER the spans, since both kinds
    ' override what the lexer decided.
    PaintTranscript shp, tr, langId, dimming, emph

    ' Bold LAST, after every colour is on. Setting Bold over a range touches
    ' only the weight - unlike Font.Highlight, which collapses the runs and
    ' repaints them one colour - but the order costs nothing and the test
    ' checks a token's colour survives it.
    BoldFocusLine shp, tr, emph

    ApplyHighlight = applied
End Function

' Repaints the two kinds of transcript line after the spans are on.
'
' Which is which is read from the TEXT, not from a stored list: a line carrying
' a prompt is something you typed, a line without one is something the
' interpreter printed. The prompt is coloured green because it is not code and
' the lexer has just coloured it as whatever ">>>" tokenizes to; the output
' loses its colour entirely, which is the signal - a printed value is not code,
' and the absence of syntax colour says so more plainly than any border.
Private Sub PaintTranscript(ByVal shp As Shape, ByVal tr As TextRange, _
                            ByVal langId As String, ByVal dimming As Boolean, _
                            ByVal emph As String)
    Dim lang As LangDef, lines() As String, i As Long, n As Long
    Dim startIdx As Long, length As Long, pl As Long
    Dim markC As Long, outC As Long, lit As Boolean

    On Error Resume Next
    If Not modOutput.IsTranscript(shp) Then
        ' Back to an editor. A block that is no longer a transcript must not
        ' keep the terminal fill, or turning it off leaves no way back.
        shp.fill.ForeColor.RGB = ThemeBackColor()
        Exit Sub
    End If
    shp.fill.ForeColor.RGB = ThemeOutputFill()

    lang = modLangRegistry.GetLang(langId)
    lines = modBlock.SplitLines(tr.text)

    For i = LBound(lines) To UBound(lines)
        n = i - LBound(lines) + 1
        If Len(Trim$(lines(i))) = 0 Then GoTo NextLine

        modBlock.LineCharRange tr.text, n, startIdx, length
        If startIdx < 1 Or length < 1 Then GoTo NextLine

        ' THE DIMMING APPLIES HERE TOO. This runs after the spans, so it would
        ' otherwise repaint the prompts and the output at full brightness and
        ' leave them lit while the code around them faded - which on a Step
        ' through means every result on the slide shouts over the one line the
        ' slide is about.
        lit = (Not dimming)
        If dimming Then
            lit = (InStr("," & emph & ",", "," & CStr(n) & ",") > 0)
        End If
        markC = ThemeOutputMark()
        outC = ThemeOutputText()
        If Not lit Then
            markC = ThemeDimmed(markC)
            outC = ThemeDimmed(outC)
        End If

        pl = modOutput.PromptLen(lines(i), lang)
        If pl > 0 Then
            tr.Characters(startIdx, pl).Font.Color.RGB = markC
        Else
            tr.Characters(startIdx, length).Font.Color.RGB = outC
            tr.Characters(startIdx, length).Font.Bold = msoFalse
        End If
NextLine:
    Next i
End Sub

' Bolds the line the slide is about: the newest emphasised one, or the one a
' walkthrough recorded when it built the slide.
'
' NOT the line an arrow points at. That was the first version and it bolded a
' line whenever an arrow was placed by hand, which is not what an arrow is for -
' an arrow points at code that stays exactly as it reads. A generated slide says
' outright which line it is about, so nothing has to be inferred.
'
' The NEWEST emphasised line, not all of them: Build up emphasises everything so
' far, and bolding the lot would say nothing about where it has got to.
Private Sub BoldFocusLine(ByVal shp As Shape, ByVal tr As TextRange, _
                          ByVal emph As String)
    Dim ln As Long, startIdx As Long, length As Long

    ln = modBlock.LastEmphasisedLine(emph)
    If ln < 1 Then ln = modBlock.GetFocusLine(shp)
    If ln < 1 Then Exit Sub

    modBlock.LineCharRange tr.text, ln, startIdx, length
    If startIdx < 1 Or length < 1 Then Exit Sub

    On Error Resume Next
    tr.Characters(startIdx, length).Font.Bold = msoTrue
    On Error GoTo 0
End Sub

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

' Draws the emphasis bands: one full-width translucent rectangle per RUN of
' consecutive emphasised lines.
'
' Rectangles rather than the text highlight, for two reasons. The text
' highlight follows the CHARACTERS, so it stops at the end of each line and
' leaves a ragged right edge, and it skips blank lines inside a range - both
' plainly visible once more than one line is emphasised. And setting
' Font.Highlight over a range COLLAPSES the runs, which repaints every
' character one colour; dropping it removes that hazard from the renderer.
'
' The rectangles sit on top, since nothing can be placed between a shape's fill
' and its own text. At this transparency the tint is slight and every token
' colour survives it - checked against a render, not assumed.
Public Sub ApplyEmphasis(ByVal shp As Shape)
    Dim spec As String, sld As Slide, lineCount As Long
    Dim size As Single, lineH As Single, pad As Single
    Dim i As Long, runStart As Long, inRun As Boolean, emphasised As Boolean

    On Error GoTo Done

    ClearBands shp
    spec = modBlock.GetEmphasis(shp)
    If Len(spec) = 0 Then Exit Sub
    If shp.Tags(modBlock.TAG_NOBAND) = "1" Then Exit Sub

    Set sld = modGutter.OwningSlide(shp)
    size = modBlock.BlockFontSize(shp)
    lineH = modSpec.SpecLine(size)
    pad = modSpec.SpecPad(size)
    lineCount = modBlock.CountLines(shp.TextFrame.TextRange.text)

    runStart = -1
    For i = 1 To lineCount + 1
        emphasised = False
        If i <= lineCount Then
            emphasised = (InStr("," & spec & ",", "," & CStr(i) & ",") > 0)
        End If

        If emphasised Then
            If runStart < 0 Then runStart = i
        ElseIf runStart >= 0 Then
            AddBand sld, shp, _
                    shp.Top + pad + (runStart - 1) * lineH, _
                    shp.Top + pad + (i - 1) * lineH
            runStart = -1
        End If
    Next i
Done:
End Sub

Public Sub ClearBands(ByVal shp As Shape)
    Dim sld As Slide, blockId As String, doomed As Collection, s2 As Shape, i As Long

    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Sub
    Set sld = modGutter.OwningSlide(shp)

    Set doomed = New Collection
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modBlock.TAG_BAND_OF) = blockId Then doomed.Add s2
    Next s2
    For i = doomed.count To 1 Step -1
        doomed(i).Delete
    Next i
End Sub

Private Sub AddBand(ByVal sld As Slide, ByVal shp As Shape, _
                    ByVal y0 As Single, ByVal y1 As Single)
    Dim r As Shape
    Set r = sld.Shapes.AddShape(msoShapeRectangle, shp.Left, y0, shp.Width, y1 - y0)
    With r
        .Fill.Solid
        .Fill.ForeColor.RGB = ThemeEmphasisColor()
        .Fill.Transparency = 0.86
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Tags.Add modBlock.TAG_BAND_OF, shp.Tags(modBlock.TAG_ID)
        .ZOrder msoBringToFront
    End With
End Sub

'==============================================================================
' Hidden lines
'==============================================================================

' Covers the hidden lines with an opaque panel carrying a question mark.
'
' The code is COVERED, not removed: the text stays in the shape, so the layout
' is unchanged, nothing can be lost, and revealing is just deleting a shape.
' The cost is that the text is still in the file - fine when the reveal happens
' in the room, not if the deck goes out beforehand.
Public Sub DrawCovers(ByVal shp As Shape)
    Dim spec As String, sld As Slide, lineCount As Long
    Dim size As Single, lineH As Single, pad As Single
    Dim i As Long, runStart As Long, hidden As Boolean

    On Error GoTo Done

    ClearCovers shp
    spec = modBlock.GetHidden(shp)
    If Len(spec) = 0 Then Exit Sub

    Set sld = modGutter.OwningSlide(shp)
    size = modBlock.BlockFontSize(shp)
    lineH = modSpec.SpecLine(size)
    pad = modSpec.SpecPad(size)
    lineCount = modBlock.CountLines(shp.TextFrame.TextRange.text)

    runStart = -1
    For i = 1 To lineCount + 1
        hidden = False
        If i <= lineCount Then
            hidden = (InStr("," & spec & ",", "," & CStr(i) & ",") > 0)
        End If

        If hidden Then
            If runStart < 0 Then runStart = i
        ElseIf runStart >= 0 Then
            AddCover sld, shp, size, _
                     CoverTop(shp, runStart, shp.Top + pad + (runStart - 1) * lineH), _
                     CoverBottom(shp, i - 1, shp.Top + pad + (i - 1) * lineH)
            runStart = -1
        End If
    Next i
Done:
End Sub

' The top and bottom of a cover, MEASURED rather than predicted.
'
' The computed line box is [pad + (n-1)*lineH, pad + n*lineH], and the glyphs do
' not stay inside it. Line spacing is set in exact points, so PowerPoint puts
' the baseline where the spacing says and lets the descenders of the last line
' hang below the box - a couple of points of code peeking out from under an
' otherwise opaque panel, which is exactly the kind of thing that only shows up
' on a projector. So ask where the characters actually landed.
'
' The measurement is unioned with the computed box rather than replacing it: a
' blank line has no characters to measure, and a failed measurement must make
' the panel bigger or leave it alone, never smaller.
Private Function CoverTop(ByVal shp As Shape, ByVal lineNo As Long, _
                          ByVal computed As Single) As Single
    Dim t As Single, h As Single
    CoverTop = computed
    If Not LineBounds(shp, lineNo, t, h) Then Exit Function
    If t < computed Then CoverTop = t
End Function

Private Function CoverBottom(ByVal shp As Shape, ByVal lineNo As Long, _
                             ByVal computed As Single) As Single
    Dim t As Single, h As Single
    CoverBottom = computed
    If Not LineBounds(shp, lineNo, t, h) Then Exit Function
    If t + h > computed Then CoverBottom = t + h
End Function

' Where one line's characters actually sit. False when there is nothing to
' measure - a blank line, or a line number outside the text.
'
' Public so a test can assert what the panel is supposed to guarantee: that no
' hidden line's ink falls outside it. Two points of leak is invisible in a
' screenshot and obvious on a projector.
Public Function LineBounds(ByVal shp As Shape, ByVal lineNo As Long, _
                            ByRef topPt As Single, ByRef heightPt As Single) As Boolean
    Dim tr As TextRange, startIdx As Long, length As Long

    On Error GoTo Done
    Set tr = shp.TextFrame.TextRange
    modBlock.LineCharRange tr.text, lineNo, startIdx, length
    If startIdx < 1 Or length < 1 Then Exit Function

    topPt = tr.Characters(startIdx, length).BoundTop
    heightPt = tr.Characters(startIdx, length).BoundHeight
    LineBounds = (heightPt > 0)
Done:
End Function

Public Sub ClearCovers(ByVal shp As Shape)
    Dim sld As Slide, blockId As String, doomed As Collection, s2 As Shape, i As Long

    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Sub
    Set sld = modGutter.OwningSlide(shp)

    Set doomed = New Collection
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modBlock.TAG_COVER_OF) = blockId Then doomed.Add s2
    Next s2
    For i = doomed.count To 1 Step -1
        doomed(i).Delete
    Next i
End Sub

Private Sub AddCover(ByVal sld As Slide, ByVal shp As Shape, ByVal size As Single, _
                     ByVal y0 As Single, ByVal y1 As Single)
    Dim r As Shape

    ' Starts at the code, not at the block edge, so line numbers stay readable -
    ' the audience can still see WHICH lines are missing.
    Set r = sld.Shapes.AddShape(msoShapeRectangle, _
                                shp.Left + shp.TextFrame.MarginLeft - size * 0.2, y0, _
                                shp.Width - shp.TextFrame.MarginLeft + size * 0.2, y1 - y0)
    With r
        .Fill.Solid
        .Fill.ForeColor.RGB = ThemeCoverColor()
        .Fill.Transparency = 0
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Tags.Add modBlock.TAG_COVER_OF, shp.Tags(modBlock.TAG_ID)
    End With

    With r.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeNone
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .TextRange.text = "?"
        .TextRange.Font.Name = THEME_FONT
        .TextRange.Font.size = size
        ' Bold and white. The question mark is the ASK - it is what the slide is
        ' doing while the code is covered - so it carries the panel rather than
        ' sitting in it quietly like a line number.
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Color.RGB = ThemeCoverMarkColor()
        .TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End With

    ' Above everything, or it would not cover anything.
    r.ZOrder msoBringToFront
End Sub
