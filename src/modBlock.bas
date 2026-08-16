Attribute VB_Name = "modBlock"
'==============================================================================
' modBlock - creating a code block, and finding one again.
'
' Creation, selection, identity and re-hugging the height. No gutter and no size
' ladder yet - those are Phases 3 and 4.
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

' Undoes what PowerPoint's autocorrect does to code as you type.
'
' Autocorrect replaces straight quotes with typographic ones, and the damage is
' invisible until the code stops being code: the lexer no longer sees a string
' at all, so it colours the contents as a variable, and pasting the text back
' into an editor is a syntax error. Repairing it on Highlight fixes the colours
' and the code together.
'
' This IS a destructive edit to the text in the shape, and a deliberate one -
' the same question PLAN.md section 14 raises about tabs. The difference is that
' a curly quote in code is never what anyone meant, whereas a tab might be.
'
' Autocapitalisation cannot be repaired here. Nothing distinguishes a variable
' the user named X from one autocorrect capitalised, so that has to be prevented
' in PowerPoint's own settings.
Public Function NormalizeCodeText(ByVal text As String) As String
    text = Replace(text, ChrW(&H2018), "'")      ' left single quote
    text = Replace(text, ChrW(&H2019), "'")      ' right single quote
    text = Replace(text, ChrW(&H201A), "'")      ' low single quote
    text = Replace(text, ChrW(&H201B), "'")      ' reversed single quote
    text = Replace(text, ChrW(&H201C), Chr$(34)) ' left double quote
    text = Replace(text, ChrW(&H201D), Chr$(34)) ' right double quote
    text = Replace(text, ChrW(&H201E), Chr$(34)) ' low double quote
    text = Replace(text, ChrW(&H2013), "-")      ' en dash, from "--"
    text = Replace(text, ChrW(&H2014), "-")      ' em dash
    text = Replace(text, ChrW(&H2026), "...")    ' ellipsis
    text = Replace(text, ChrW(&HA0), " ")        ' non-breaking space
    NormalizeCodeText = text
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

    ' Tab stops every four characters, matching an editor. Without this a tab
    ' jumps to PowerPoint's default inch stop, so an indented line sits far
    ' right of where the same code sits in VS Code.
    On Error Resume Next
    shp.TextFrame.Ruler.TabStops.DefaultSpacing = modSpec.SpecTabStop(size)
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Finding a block again
'------------------------------------------------------------------------------

' The block implied by the current selection, or Nothing with a reason in
' problem. Deliberately permissive: PLAN.md section 8 says any selected shape
' with text counts, tagged or not, so a block pasted from an old deck or from a
' colleague still works. Nothing about DISPLAY depends on the tag surviving.
Public Function SelectedBlock(ByRef problem As String) As Shape
    Dim sel As Selection, shp As Shape

    problem = ""
    On Error GoTo NoWindow
    Set sel = Application.ActiveWindow.Selection
    On Error GoTo 0

    Select Case sel.Type
        Case ppSelectionShapes, ppSelectionText
            If sel.ShapeRange.count = 0 Then
                problem = "Select a code block first."
                Exit Function
            End If
            Set shp = sel.ShapeRange(1)
        Case Else
            problem = "Select a code block first."
            Exit Function
    End Select

    If shp.Type = msoGroup Then
        problem = "That is a group. Ungroup it, or select the block inside it."
        Exit Function
    End If
    If shp.HasTextFrame = msoFalse Then
        problem = "That shape cannot hold text, so there is nothing to highlight."
        Exit Function
    End If
    If shp.TextFrame.HasText = msoFalse Then
        problem = "That block is empty. Type some code into it first."
        Exit Function
    End If

    Set SelectedBlock = shp
    Exit Function

NoWindow:
    problem = "Open a slide in Normal view first."
End Function

Public Function IsCodeBlock(ByVal shp As Shape) As Boolean
    IsCodeBlock = (shp.Tags(TAG_BLOCK) = "1")
End Function

' Adds the bookkeeping tags if they are absent, leaving any that are already
' there alone. Called on every Highlight, which is what adopts an untagged
' block the first time someone highlights it.
Public Sub EnsureTags(ByVal shp As Shape, ByVal langId As String)
    If Len(shp.Tags(TAG_BLOCK)) = 0 Then shp.Tags.Add TAG_BLOCK, "1"
    If Len(shp.Tags(TAG_ID)) = 0 Then shp.Tags.Add TAG_ID, NewBlockId()
    If Len(shp.Tags(TAG_LANG)) = 0 Then shp.Tags.Add TAG_LANG, langId
End Sub

' Resolution order: the shape's own tag, then the ribbon's current choice, then
' the registry default. An unrecognised tag falls back rather than erroring, so
' a block from a future version still renders as something.
Public Function BlockLangId(ByVal shp As Shape, ByVal fallbackId As String) As String
    Dim v As String
    v = shp.Tags(TAG_LANG)
    If Len(v) = 0 Then v = fallbackId
    If Len(v) = 0 Then v = modLangRegistry.DefaultLangId()
    BlockLangId = v
End Function

Public Sub SetBlockLang(ByVal shp As Shape, ByVal langId As String)
    shp.Tags.Add TAG_LANG, langId      ' Add replaces an existing tag
End Sub

' The size the block is currently set in. Reading the first character rather
' than the whole range, because a range with mixed sizes reports ppMixed and
' the caller would silently get nonsense.
Public Function BlockFontSize(ByVal shp As Shape) As Single
    Dim s As Single
    On Error Resume Next
    s = shp.TextFrame.TextRange.Characters(1, 1).Font.size
    On Error GoTo 0
    If s <= 0 Then s = modSpec.BASE_SIZE
    BlockFontSize = s
End Function

'------------------------------------------------------------------------------

' Re-hugs the block to its content. Autofit is off, so a block does not grow as
' you type - this is what makes the flow work: insert, type, press Highlight,
' and the block fits again. It is also the recovery path if anything drifts.
Public Sub ResizeToContent(ByVal shp As Shape)
    Dim size As Single, h As Single, pad As Single, centreY As Single

    size = BlockFontSize(shp)
    h = modSpec.SpecHeight(size, CountLines(shp.TextFrame.TextRange.text))
    pad = modSpec.SpecPad(size)

    With shp.TextFrame
        .MarginLeft = pad
        .MarginRight = pad
        .MarginTop = pad
        .MarginBottom = pad
    End With

    ' Grow around the block's CENTRE, not its top edge. Growing downward from a
    ' fixed top sends a block that started one line tall off the bottom of the
    ' slide as soon as real code is typed into it - and the height is correct
    ' the whole time, so only looking at a render catches it.
    '
    ' Preserving the centre also does the right thing for a block the user has
    ' deliberately moved: it stays put instead of jumping back to the middle.
    ' And a block created centred in the content area stays exactly where the
    ' reference puts it.
    centreY = shp.Top + shp.Height / 2
    shp.Height = h
    shp.Top = centreY - h / 2

    ' Only clamp when the block can actually fit. When it cannot, leave it
    ' centred and overflowing equally at both ends - that is the honest signal
    ' that the content needs a smaller size, which is what Fit is for.
    If h <= modSpec.SLIDE_H Then
        If shp.Top < 0 Then shp.Top = 0
        If shp.Top + h > modSpec.SLIDE_H Then shp.Top = modSpec.SLIDE_H - h
    End If

    shp.Adjustments(1) = modSpec.SpecCornerAdjust(size, h)
End Sub

Private Function NewBlockId() As String
    ' Good enough to tell two blocks on a slide apart, which is all the gutter
    ' needs. Not a guid, and does not need to be.
    NewBlockId = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 100000))
End Function
