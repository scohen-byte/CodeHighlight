Attribute VB_Name = "modBlock"
'==============================================================================
' modBlock - creating a code block, and finding one again.
'
' THIN SLICE VERSION: creation only. No gutter, no resizing, no identity
' fallback. Those come after the render path is proven against a real slide.
'
' The shape's text is the source. Nothing is stored anywhere else, which is what
' makes the output portable - a colleague without the add-in sees an ordinary
' rounded rectangle with ordinary per-character font colours, and can edit it.
'==============================================================================
Option Explicit

' Shape tags. Language-neutral names, because they persist inside decks that
' have already been handed out - renaming them later orphans every block in the
' wild. See PLAN.md section 5b.
Public Const TAG_BLOCK As String = "CODEBLOCK"
Public Const TAG_ID    As String = "CODEBLOCK_ID"
Public Const TAG_LANG  As String = "CODEBLOCK_LANG"

' PowerPoint separates paragraphs with CR. Text arriving from a file has LF or
' CRLF, and assigning that directly produces either one giant paragraph or
' stray vertical tabs, both of which wreck the line count and the height.
' Trailing blank lines are stripped. A file almost always ends with a newline,
' and keeping it would add an empty paragraph - so the block would stand one
' line taller than its content and no longer hug it. tools/lab.py drops trailing
' empties the same way, and the difference is plainly visible side by side.
Public Function NormalizeParagraphs(ByVal text As String) As String
    text = Replace(text, vbCrLf, vbCr)
    text = Replace(text, vbLf, vbCr)
    Do While Len(text) > 0
        If Right$(text, 1) <> vbCr Then Exit Do
        text = Left$(text, Len(text) - 1)
    Loop
    NormalizeParagraphs = text
End Function

Public Function CountLines(ByVal text As String) As Long
    Dim n As Long, i As Long
    n = 1
    For i = 1 To Len(text)
        If Mid$(text, i, 1) = vbCr Then n = n + 1
    Next i
    CountLines = n
End Function

' Inserts a code block on sld and returns it. Geometry comes entirely from
' modSpec, so the only thing that varies is the size.
Public Function CreateBlock(ByVal sld As Slide, ByVal code As String, _
                            ByVal size As Single, ByVal langId As String) As Shape
    Dim shp As Shape
    Dim lineCount As Long, h As Single, pad As Single

    code = NormalizeParagraphs(code)
    lineCount = CountLines(code)
    h = modSpec.SpecHeight(size, lineCount)
    pad = modSpec.SpecPad(size)

    ' Centred in the content area, the way the lab renders it.
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  modSpec.CONTENT_L, _
                                  modSpec.CONTENT_T + (modSpec.CONTENT_H - h) / 2, _
                                  modSpec.CONTENT_W, h)

    With shp
        .Fill.Solid
        .Fill.ForeColor.RGB = ThemeBackColor()
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Adjustments(1) = modSpec.SpecCornerAdjust(size, h)
    End With

    With shp.TextFrame
        ' Wrap and autofit both off. Either one would misalign the gutter later,
        ' and wrapped code is unreadable anyway.
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeNone
        .VerticalAnchor = msoAnchorTop
        .MarginLeft = pad
        .MarginRight = pad
        .MarginTop = pad
        .MarginBottom = pad
        .TextRange.text = code
    End With

    FormatBlockText shp, size

    shp.Tags.Add TAG_BLOCK, "1"
    shp.Tags.Add TAG_ID, NewBlockId()
    shp.Tags.Add TAG_LANG, langId

    Set CreateBlock = shp
End Function

' Font, alignment and paragraph spacing. Separated out because resizing has to
' reapply exactly this set, and missing one of them is what makes a resized
' block look subtly wrong.
Public Sub FormatBlockText(ByVal shp As Shape, ByVal size As Single)
    With shp.TextFrame.TextRange
        .Font.Name = THEME_FONT
        .Font.size = size
        .Font.Color.RGB = ThemeColor(tkDefault)

        With .ParagraphFormat
            ' Autoshapes default to CENTRED text. Code must be explicitly left
            ' aligned or the whole block renders ragged, and it is invisible
            ' until you look at a render.
            .Alignment = ppAlignLeft
            ' LineRuleWithin False means SpaceWithin is read as points rather
            ' than as a multiple of the line height.
            .LineRuleWithin = msoFalse
            .SpaceWithin = modSpec.SpecLine(size)
            .LineRuleBefore = msoFalse
            .SpaceBefore = 0
            .LineRuleAfter = msoFalse
            .SpaceAfter = 0
        End With
    End With
End Sub

Private Function NewBlockId() As String
    ' Good enough to tell two blocks on a slide apart, which is all the gutter
    ' needs. Not a guid, and does not need to be.
    NewBlockId = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 100000))
End Function
