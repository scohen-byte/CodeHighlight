Attribute VB_Name = "modRender"
'==============================================================================
' modRender - applies the scanner's spans to a shape.
'
' The whole module is about one loop, and two things make it correct:
'
'   RESET FIRST, THEN COLOUR. Setting the entire range to the default colour
'   before applying spans is what makes re-highlighting idempotent. Without it,
'   editing a line and pressing Highlight again leaves the old colours behind
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

    Set tr = shp.TextFrame.TextRange
    If Len(tr.text) = 0 Then Exit Function

    tr.Font.Name = THEME_FONT
    tr.Font.Color.RGB = ThemeColor(tkDefault)

    n = modLexer.Tokenize(tr.text, modLangRegistry.GetLang(langId), spans)

    For i = 0 To n - 1
        ' Default-coloured spans were already handled by the reset above, and
        ' skipping them removes roughly half the round trips.
        If spans(i).Kind <> tkDefault Then
            tr.Characters(spans(i).Start, spans(i).Length).Font.Color.RGB = _
                ThemeColor(spans(i).Kind)
            applied = applied + 1
        End If
    Next i

    ApplyHighlight = applied
End Function
