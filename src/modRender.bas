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
