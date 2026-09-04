Attribute VB_Name = "modRetheme"
'==============================================================================
' modRetheme - moves a block from one theme to the other.
'
' THE RULE: the theme owns COLOUR, the user owns everything else.
'
' Nothing here changes a size, a weight, a font, a position or a shape's
' geometry, and nothing here creates or deletes a shape. That is not tidiness,
' it is the whole design. Stylize is a re-render - it refits the block, rebuilds
' the bands, and re-derives every note's size from the block - so running it to
' change a palette silently throws away hand-made adjustments. Doing that once
' is annoying. Doing it across 123 blocks because somebody picked a theme from a
' dropdown is not recoverable.
'
' modNote's own comment for note colour already states the principle: "Colour
' changes no geometry, so there is nothing to re-place." This is that, one level
' up.
'
' WHAT IS DELIBERATELY LEFT ALONE
'   Notes      - a note's colour is a choice the user made about that note. The
'                theme's only say is which preset a NEW note starts as.
'   Arrows     - drawn on the slide, not on the block, so the block's ground is
'                not the background they have to survive.
'   Guides     - furniture with its own colour, handled by modGuides.
'==============================================================================
Option Explicit

' Applies a theme to one block and everything hanging off it.
'
' Returns the number of colour writes made, which is the number of COM round
' trips and therefore the only number worth reporting. Measured in-process a
' write costs about 1.6 ms, so the count is the runtime: a real 84-slide deck
' needs 2,773 of them and takes 4.4 seconds.
'
' Driving the same writes from OUTSIDE PowerPoint measured 8.5 ms each, five
' times worse. Worth knowing before optimising against a number taken from a
' harness: the add-in does not pay that, because it runs in-process.
Public Function RethemeBlock(ByVal shp As Shape, ByVal t As ThemeId) As Long
    Dim langId As String, writes As Long

    On Error GoTo Failed

    ' Already there. Worth checking rather than repainting: it takes a second
    ' press of "apply to all" from 4.4 seconds to nothing, and pressing it again
    ' is exactly what somebody does when they are not sure it worked.
    If modBlock.BlockTheme(shp) = t And Len(shp.Tags(modBlock.TAG_THEME)) > 0 Then
        Exit Function
    End If

    modBlock.SetBlockTheme shp, t
    ThemeSetCurrent t

    langId = modBlock.BlockLangId(shp, "")
    writes = modRender.RecolorText(shp, langId)

    ' PaintTranscript inside RecolorText already set the block fill and border
    ' for a transcript. A plain code block still needs both.
    If Not modOutput.IsTranscript(shp) Then
        shp.fill.ForeColor.RGB = ThemeBackColor()
        modBlock.ApplyBlockEdge shp
        writes = writes + 2
    End If

    writes = writes + RethemeParts(shp)
    RethemeBlock = writes
    Exit Function
Failed:
    ' A block that fails is skipped, not fatal. One damaged shape must not stop
    ' a deck-wide pass part way through and leave half a deck in each theme.
    RethemeBlock = writes
End Function

' The shapes that belong to a block: bands, covers, the gutter, leaders and
' output notes. Found the same way GroupParts finds them, by the *_OF tags.
'
' Prose notes are matched and skipped rather than not matched, so that the
' decision is visible here rather than being an accident of which tags this
' happens to test.
Private Function RethemeParts(ByVal shp As Shape) As Long
    Dim sld As Slide, blockId As String, writes As Long
    Dim p As Shape

    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Function
    Set sld = modGutter.OwningSlide(shp)
    If sld Is Nothing Then Exit Function

    ' AllShapes, not Shapes. A styled block is GROUPED with its gutter, guides,
    ' bands and notes, and a group hides every one of them from sld.Shapes - so
    ' walking the slide directly finds the parts of unstyled blocks only. On a
    ' real deck that silently skipped 56 of 123 blocks and all of their parts.
    For Each p In modBlock.AllShapes(sld)

        If p.Tags(modBlock.TAG_BAND_OF) = blockId Then
            p.fill.ForeColor.RGB = ThemeEmphasisColor()
            writes = writes + 1

        ElseIf p.Tags(modBlock.TAG_COVER_OF) = blockId Then
            p.fill.ForeColor.RGB = ThemeCoverColor()
            writes = writes + 1
            On Error Resume Next
            p.TextFrame.TextRange.Font.Color.RGB = ThemeCoverMarkColor()
            On Error GoTo 0
            writes = writes + 1

        ElseIf p.Tags(modGutter.TAG_GUTTER_OF) = blockId Then
            On Error Resume Next
            p.TextFrame.TextRange.Font.Color.RGB = ThemeGutterColor()
            On Error GoTo 0
            writes = writes + 1

        ElseIf p.Tags(modNote.TAG_LEADER_OF) = blockId Then
            p.Line.ForeColor.RGB = ThemeLeaderColor()
            writes = writes + 1

        ElseIf p.Tags(modNote.TAG_NOTE_OF) = blockId Then
            ' An OUTPUT note is part of the code - it is the terminal, and it
            ' has no colour of its own that the user chose. A prose note does,
            ' so it is left exactly as it is.
            If p.Tags(modNote.TAG_NOTE_KIND) = modNote.KIND_OUTPUT Then
                writes = writes + RethemeOutputNote(p)
            End If
        End If
    Next p

    RethemeParts = writes
End Function

' An output note carries four colours and they have to move together. Getting
' this wrong is not a cosmetic bug: flipping the fill while leaving the text
' behind puts light grey on light grey, and the line vanishes completely.
Private Function RethemeOutputNote(ByVal note As Shape) As Long
    Dim tr As TextRange, markLen As Long

    On Error Resume Next
    note.fill.ForeColor.RGB = ThemeOutputFill()
    note.Line.ForeColor.RGB = ThemeOutputEdge()

    Set tr = note.TextFrame.TextRange
    If tr Is Nothing Then
        RethemeOutputNote = 2
        Exit Function
    End If

    tr.Font.Color.RGB = ThemeOutputText()

    ' The arrow that opens the line keeps its own colour. It is the one thing on
    ' an output note that says "this is not the code", so it is coloured last
    ' and separately - the same reason ThemeOutputMark is stated outright rather
    ' than read from the comment colour.
    markLen = Len(modNote.OutputMark())
    If markLen > 0 And Len(tr.text) >= markLen Then
        tr.Characters(1, markLen).Font.Color.RGB = ThemeOutputMark()
    End If

    RethemeOutputNote = 4
End Function

'------------------------------------------------------------------------------
' Deck-wide
'------------------------------------------------------------------------------
' Every block in the presentation. Returns the number of blocks actually
' changed; writes comes back through the ByRef so the caller can report both.
'
' Blocks already in the target theme cost nothing, so the count reported to the
' user is the count that will actually be touched.
Public Function RethemeAll(ByVal t As ThemeId, ByRef writes As Long) As Long
    Dim sld As Slide, changed As Long, w As Long
    Dim p As Presentation, shp As Shape

    On Error Resume Next
    Set p = Application.ActivePresentation
    On Error GoTo 0
    If p Is Nothing Then Exit Function

    For Each sld In p.Slides
        ' AllShapes descends into groups, which is where most blocks live once
        ' they have been styled. It also returns a snapshot rather than a live
        ' collection, so nothing here depends on the slide not changing under it.
        For Each shp In modBlock.AllShapes(sld)
            If shp.Tags(modBlock.TAG_BLOCK) = "1" Then
                w = RethemeBlock(shp, t)
                If w > 0 Then changed = changed + 1
                writes = writes + w
            End If
        Next shp
    Next sld

    RethemeAll = changed
End Function

' How many blocks a deck-wide change would actually touch, for the confirm
' prompt. Counted before anything is written, because "this will change 118
' blocks and take about five seconds" is the sentence that makes the wait
' make sense - and a count of zero means there is nothing to confirm at all.
Public Function CountBlocksToChange(ByVal t As ThemeId) As Long
    Dim sld As Slide, n As Long, p As Presentation
    Dim shp As Shape

    On Error Resume Next
    Set p = Application.ActivePresentation
    On Error GoTo 0
    If p Is Nothing Then Exit Function

    For Each sld In p.Slides
        For Each shp In modBlock.AllShapes(sld)
            If shp.Tags(modBlock.TAG_BLOCK) = "1" Then
                If modBlock.BlockTheme(shp) <> t Or _
                   Len(shp.Tags(modBlock.TAG_THEME)) = 0 Then
                    n = n + 1
                End If
            End If
        Next shp
    Next sld

    CountBlocksToChange = n
End Function
