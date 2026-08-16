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

' Why the last GroupParts did nothing, for the test harness to report.
Public LastGroupError As String

' Marks the group holding a block and its parts.
Public Const TAG_GROUP_OF As String = "CODEBLOCK_GROUP_OF"

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

' Paragraph count, matching what PowerPoint itself reports.
'
' A TRAILING paragraph mark does not start a new paragraph: PowerPoint counts
' "a" & vbCr as one paragraph, not two. Counting it as two put a phantom number
' at the bottom of the gutter, and made the numbers appear correct only when
' there happened to be a blank line at the end.
Public Function CountLines(ByVal text As String) As Long
    Dim n As Long, i As Long

    Do While Len(text) > 0
        If Right$(text, 1) <> vbCr Then Exit Do
        text = Left$(text, Len(text) - 1)
    Loop

    n = 1
    For i = 1 To Len(text)
        If Mid$(text, i, 1) = vbCr Then n = n + 1
    Next i
    CountLines = n
End Function

' The longest line, in COLUMNS not characters, so a tab counts as a full tab
' stop. This is what decides whether the code fits the slide widthwise.
Public Function LongestLine(ByVal text As String) As Long
    Dim lines() As String, i As Long, n As Long, w As Long
    text = Replace(NormalizeParagraphs(text), vbTab, Space$(CLng(modSpec.TAB_CHARS)))
    lines = Split(text, vbCr)
    For i = LBound(lines) To UBound(lines)
        w = Len(lines(i))
        If w > n Then n = w
    Next i
    LongestLine = n
End Function

' Moves a block to a new size. Resizing has to reapply EVERY derived quantity
' together - font size, exact line spacing, all four margins, tab stops, the
' height and the corner radius - and missing one is what makes a resized block
' look subtly wrong. FormatBlockText owns the text properties, ResizeToContent
' the geometry, so between them nothing is left behind.
Public Sub ApplySize(ByVal shp As Shape, ByVal newSize As Single)
    FormatBlockText shp, newSize
    ResizeToContent shp
End Sub

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

    ' Placed at the top-left of the content area, and below anything already
    ' there, rather than centred. Centring looks tidy for one block and is
    ' useless the moment you want two, or want to put one somewhere specific -
    ' every new block landed on top of the last. The block is selected on
    ' creation, so dragging it is one gesture.
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, _
                                  modSpec.CONTENT_L, NextFreeTop(sld), _
                                  modSpec.CONTENT_W, h)

    With shp
        .Fill.Solid
        .Fill.ForeColor.RGB = ThemeBackColor()
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
    End With

    With shp.TextFrame
        ' WRAP STAYS OFF. That is the property the line-number gutter depends on:
        ' one paragraph must be one visual line, or every number below a wrapped
        ' line drifts. Autofit is safe precisely because wrap is off.
        '
        ' Autofit ON so the block grows as you type instead of waiting for the
        ' next Highlight. With wrap off it grows in BOTH directions, so the block
        ' hugs its code rather than spanning the slide - which is what you want
        ' for a short snippet.
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeShapeToFitText
        .VerticalAnchor = msoAnchorTop
        .MarginLeft = pad
        .MarginRight = pad
        .MarginTop = pad
        .MarginBottom = pad
        .TextRange.text = code
    End With

    FormatBlockText shp, size

    ' Autofit has sized the shape by now, so position it from what it actually
    ' became rather than from the predicted height.
    shp.Left = modSpec.CONTENT_L
    shp.Top = modSpec.CONTENT_T + (modSpec.CONTENT_H - shp.Height) / 2
    shp.Adjustments(1) = modSpec.SpecCornerAdjust(size, ShorterSide(shp))

    shp.Tags.Add TAG_BLOCK, "1"
    shp.Tags.Add TAG_ID, NewBlockId()
    shp.Tags.Add TAG_LANG, langId

    Set CreateBlock = shp
End Function

' Just below the lowest code block on the slide, or the top of the content area
' when there is none.
Private Function NextFreeTop(ByVal sld As Slide) As Single
    Dim i As Long, lowest As Single, bottom As Single

    Dim shp As Shape
    lowest = modSpec.CONTENT_T
    For Each shp In AllShapes(sld)
        If shp.Tags(TAG_BLOCK) = "1" Then
            bottom = shp.Top + shp.Height + 12
            If bottom > lowest Then lowest = bottom
        End If
    Next shp
    If lowest > modSpec.SLIDE_H - 60 Then lowest = modSpec.CONTENT_T
    NextFreeTop = lowest
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
        ' A group holding one of our blocks is the normal case now, not an
        ' error: Highlight groups the block with its numbers and guides so the
        ' whole thing can be dragged as one.
        Dim inner As Shape
        Set inner = BlockInGroup(shp)
        If inner Is Nothing Then
            problem = "That is a group with no code block in it."
            Exit Function
        End If
        Set SelectedBlock = inner
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

'------------------------------------------------------------------------------
' Grouping
'------------------------------------------------------------------------------
' PLAN.md section 9 says not to group, on the grounds that it makes editing
' awkward. In use it turned out the opposite way round: without a group, moving
' a block means selecting three or more shapes and dragging them together, every
' time. Text is still editable - clicking twice enters the shape inside a group.
'
' Highlight ungroups, does its work, and regroups, so nothing downstream has to
' understand groups. SelectedBlock looks inside a selected group, so pressing
' Highlight with the group selected does the right thing.

' Every shape on the slide, descending INTO groups.
'
' Needed because grouping hides the parts: a gutter or a guide inside a group is
' not in sld.Shapes any more, so anything that finds parts by tag has to look
' inside groups or it will conclude they were deleted.
Public Function AllShapes(ByVal sld As Slide) As Collection
    Dim c As Collection, i As Long
    Set c = New Collection
    For i = 1 To sld.Shapes.count
        CollectDeep c, sld.Shapes(i)
    Next i
    Set AllShapes = c
End Function

Private Sub CollectDeep(ByVal c As Collection, ByVal shp As Shape)
    Dim i As Long
    c.Add shp
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.count
            CollectDeep c, shp.GroupItems(i)
        Next i
    End If
End Sub

Public Function ParentGroup(ByVal shp As Shape) As Shape
    On Error Resume Next
    Set ParentGroup = shp.ParentGroup
    On Error GoTo 0
End Function

' The code block inside a group, or Nothing.
Public Function BlockInGroup(ByVal grp As Shape) As Shape
    Dim i As Long
    On Error Resume Next
    For i = 1 To grp.GroupItems.count
        If grp.GroupItems(i).Tags(TAG_BLOCK) = "1" Then
            Set BlockInGroup = grp.GroupItems(i)
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

' Breaks the block out of its group so the parts can be redrawn. Returns True
' if it was grouped, so the caller knows to put it back.
Public Function UngroupParts(ByVal shp As Shape) As Boolean
    Dim g As Shape
    Set g = ParentGroup(shp)
    If g Is Nothing Then Exit Function
    g.Ungroup
    UngroupParts = True
End Function

' Groups the block with its gutter and guides. Does nothing when there are no
' parts to group, so a bare block stays a plain shape.
Public Sub GroupParts(ByVal shp As Shape)
    Dim sld As Slide, i As Long, blockId As String
    Dim v() As Variant, n As Long, grp As Shape

    LastGroupError = ""
    blockId = shp.Tags(TAG_ID)
    If Len(blockId) = 0 Then
        LastGroupError = "block has no id"
        Exit Sub
    End If
    Set sld = shp.Parent

    ' Shapes.Range wants a ONE-BASED Variant array. Handed a zero-based String
    ' array inside a Variant it neither groups nor raises - it just quietly does
    ' nothing, which is the worst of both.
    ReDim v(1 To sld.Shapes.count)
    For i = 1 To sld.Shapes.count
        With sld.Shapes(i)
            If .Tags(TAG_ID) = blockId Or _
               .Tags(modGutter.TAG_GUTTER_OF) = blockId Or _
               .Tags(modGuides.TAG_GUIDE_OF) = blockId Then
                n = n + 1
                v(n) = .Name
            End If
        End With
    Next i

    If n < 2 Then
        LastGroupError = "only " & n & " shape(s) to group"
        Exit Sub
    End If
    ReDim Preserve v(1 To n)

    On Error Resume Next
    Set grp = sld.Shapes.Range(v).Group
    If Err.Number <> 0 Then
        LastGroupError = "Group failed: " & Err.Description
        Err.Clear
    ElseIf grp Is Nothing Then
        LastGroupError = "Group returned nothing for " & n & " shapes"
    Else
        grp.Tags.Add TAG_GROUP_OF, blockId
    End If
    On Error GoTo 0
End Sub

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
' Autofit already keeps the block hugging its code, so this no longer sets the
' height - PowerPoint owns that now, and its computed height is a few points
' tighter than SpecHeight because it does not add full line spacing below the
' last line. What this still owns is the margins, the position and the radius.
Public Sub ResizeToContent(ByVal shp As Shape)
    Dim size As Single, h As Single, pad As Single, centreY As Single

    size = BlockFontSize(shp)
    pad = modSpec.SpecPad(size)
    centreY = shp.Top + shp.Height / 2

    With shp.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeShapeToFitText
        .MarginLeft = pad
        .MarginRight = pad
        .MarginTop = pad
        .MarginBottom = pad
    End With
    h = shp.Height

    ' Re-centre on the block's own centre rather than the slide's. Autofit grows
    ' downward from the top, so a block that started short would otherwise creep
    ' off the bottom as lines are added - and its height is correct the whole
    ' time, so only a render shows it. Preserving the centre also leaves a block
    ' the user has deliberately moved where they put it.
    shp.Top = centreY - h / 2

    ' Only clamp when the block can actually fit. When it cannot, leave it
    ' centred and overflowing equally at both ends - that is the honest signal
    ' that the content needs a smaller size, which is what Fit is for.
    If h <= modSpec.SLIDE_H Then
        If shp.Top < 0 Then shp.Top = 0
        If shp.Top + h > modSpec.SLIDE_H Then shp.Top = modSpec.SLIDE_H - h
    End If

    shp.Adjustments(1) = modSpec.SpecCornerAdjust(size, ShorterSide(shp))
End Sub

Private Function ShorterSide(ByVal shp As Shape) As Single
    ShorterSide = shp.Height
    If shp.Width < ShorterSide Then ShorterSide = shp.Width
End Function

Private Function NewBlockId() As String
    ' Good enough to tell two blocks on a slide apart, which is all the gutter
    ' needs. Not a guid, and does not need to be.
    NewBlockId = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 100000))
End Function
