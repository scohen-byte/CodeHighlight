Attribute VB_Name = "modGuides"
'==============================================================================
' modGuides - indent guides, the vertical lines marking each level of nesting.
'
' Drawn as real line shapes rather than as characters in the text, because the
' text of the block must stay pure source: it is what gets copied out, and what
' the lexer reads. The lines live in the indentation whitespace, so they never
' cross a glyph.
'
' Geometry rests on one property of the block: line spacing is set in EXACT
' points (see modSpec), so line N's box starts at MarginTop + N * SpecLine. That
' is what makes a purely geometric overlay line up with text at all, and it is
' the same reason the gutter will be able to align.
'
' Guides are tagged with the id of the block they belong to, and cleared and
' redrawn on every Highlight. That is also the recovery path: if a block is
' moved and its guides are left behind, pressing Highlight puts them right.
'==============================================================================
Option Explicit

Public Const TAG_GUIDE_OF As String = "CODEBLOCK_GUIDE_OF"
' Guides are on unless a block says otherwise, so an untagged block - one
' pasted in from elsewhere - gets them.
Public Const TAG_GUIDES_OFF As String = "CODEBLOCK_GUIDES_OFF"


' VS Code Dark+ draws indent guides in 404040 - dark enough to recede, light
' enough to follow. Deliberately NOT in modTheme's token palette: this is
' furniture, not a token class.
Private Const GUIDE_R As Long = 64
Private Const GUIDE_G As Long = 64
Private Const GUIDE_B As Long = 64
Private Const GUIDE_WEIGHT As Single = 0.75

' ALL module-level declarations must come before ANY procedure. Putting these
' two functions above the constants left GUIDE_R and friends stranded after a
' procedure, and VBA reports that as "Compile error: Variable not defined" in
' AddGuideLine - naming the USE site, not the declaration that is misplaced.
Public Function GuidesEnabled(ByVal shp As Shape) As Boolean
    GuidesEnabled = (shp.Tags(TAG_GUIDES_OFF) <> "1")
End Function

Public Sub SetGuidesEnabled(ByVal shp As Shape, ByVal enabled As Boolean)
    shp.Tags.Add TAG_GUIDES_OFF, IIf(enabled, "0", "1")
End Sub

' Removes every guide belonging to a block. Iterates backwards, because
' deleting from a Shapes collection while walking it forwards skips shapes.
Public Sub ClearGuides(ByVal sld As Slide, ByVal blockId As String)
    Dim doomed As Collection, shp As Shape, i As Long

    If Len(blockId) = 0 Then Exit Sub      ' "" would match every untagged shape

    ' Collect first, delete after. Deleting while walking a collection that
    ' descends into groups is a good way to skip shapes or trip over a group
    ' that dissolves when its second-to-last child goes.
    Set doomed = New Collection
    For Each shp In modBlock.AllShapes(sld)
        If shp.Tags(TAG_GUIDE_OF) = blockId Then doomed.Add shp
    Next shp
    For i = doomed.count To 1 Step -1
        doomed(i).Delete
    Next i
End Sub

' Draws the guides for one block, returning how many line shapes were created.
Public Function DrawGuides(ByVal shp As Shape) As Long
    Dim sld As Slide, blockId As String
    Dim levels() As Long, lineCount As Long
    Dim size As Single, charW As Single, lineH As Single, pad As Single
    Dim col As Long, i As Long, runStart As Long, drawn As Long
    Dim inRun As Boolean
    Dim x As Single, y0 As Single, y1 As Single
    Dim maxLevel As Long

    Set sld = modGutter.OwningSlide(shp)
    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Function

    ClearGuides sld, blockId
    If Not GuidesEnabled(shp) Then Exit Function

    lineCount = IndentLevels(shp.TextFrame.TextRange.text, levels)
    If lineCount = 0 Then Exit Function

    size = modBlock.BlockFontSize(shp)
    charW = modSpec.SpecCharW(size)
    lineH = modSpec.SpecLine(size)
    pad = modSpec.SpecPad(size)

    For i = 0 To lineCount - 1
        If levels(i) > maxLevel Then maxLevel = levels(i)
    Next i

    ' One pass per indent column. A guide at column C is present on any line
    ' indented deeper than C, so consecutive such lines merge into a single
    ' shape instead of one per line.
    For col = 0 To maxLevel - 1
        ' The block's ACTUAL left margin, not the padding: with a line-number
        ' gutter the text starts further right, and guides must follow it.
        x = shp.Left + shp.TextFrame.MarginLeft + col * modSpec.TAB_CHARS * charW
        runStart = -1
        ' Runs to lineCount INCLUSIVE, so a run reaching the last line still
        ' gets closed. The guard has to be nested rather than written as
        ' "i < lineCount And levels(i) > col": VBA does NOT short-circuit And,
        ' so it would evaluate levels(lineCount) and raise subscript out of
        ' range on the final pass.
        For i = 0 To lineCount
            inRun = False
            If i < lineCount Then
                If levels(i) > col Then inRun = True
            End If

            If inRun Then
                If runStart < 0 Then runStart = i
            ElseIf runStart >= 0 Then
                y0 = shp.Top + pad + runStart * lineH
                y1 = shp.Top + pad + i * lineH
                AddGuideLine sld, blockId, x, y0, y1
                drawn = drawn + 1
                runStart = -1
            End If
        Next i
    Next col

    DrawGuides = drawn
End Function

Private Sub AddGuideLine(ByVal sld As Slide, ByVal blockId As String, _
                         ByVal x As Single, ByVal y0 As Single, ByVal y1 As Single)
    Dim ln As Shape
    Set ln = sld.Shapes.AddLine(x, y0, x, y1)
    With ln.Line
        .ForeColor.RGB = RGB(GUIDE_R, GUIDE_G, GUIDE_B)
        .Weight = GUIDE_WEIGHT
    End With
    ln.Tags.Add TAG_GUIDE_OF, blockId
End Sub

'------------------------------------------------------------------------------

' Indent level per line, in units of TAB_CHARS columns. Returns the line count.
'
' A blank line takes the SMALLER of its neighbours' levels, which is what an
' editor does: a blank line inside a loop body keeps the guide running through
' it, but a blank line between two functions does not sprout one.
Public Function IndentLevels(ByVal text As String, ByRef levels() As Long) As Long
    Dim lines() As String, n As Long, i As Long, j As Long
    Dim cols() As Long, blank() As Boolean
    Dim prevLevel As Long, nextLevel As Long

    text = Replace(Replace(text, vbCrLf, vbCr), vbLf, vbCr)
    lines = Split(text, vbCr)
    n = UBound(lines) - LBound(lines) + 1
    If n <= 0 Then Exit Function

    ReDim levels(0 To n - 1)
    ReDim cols(0 To n - 1)
    ReDim blank(0 To n - 1)

    For i = 0 To n - 1
        blank(i) = (Len(Trim$(lines(i))) = 0)
        cols(i) = LeadingColumns(lines(i))
    Next i

    For i = 0 To n - 1
        If Not blank(i) Then
            levels(i) = cols(i) \ CLng(modSpec.TAB_CHARS)
        Else
            prevLevel = 0
            For j = i - 1 To 0 Step -1
                If Not blank(j) Then prevLevel = cols(j) \ CLng(modSpec.TAB_CHARS): Exit For
            Next j
            nextLevel = 0
            For j = i + 1 To n - 1
                If Not blank(j) Then nextLevel = cols(j) \ CLng(modSpec.TAB_CHARS): Exit For
            Next j
            levels(i) = IIf(prevLevel < nextLevel, prevLevel, nextLevel)
        End If
    Next i

    IndentLevels = n
End Function

' Leading whitespace measured in COLUMNS, so a tab counts as a full tab stop
' rather than as one character.
Private Function LeadingColumns(ByVal line As String) As Long
    Dim i As Long, ch As String, c As Long
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = " " Then
            c = c + 1
        ElseIf ch = vbTab Then
            c = ((c \ CLng(modSpec.TAB_CHARS)) + 1) * CLng(modSpec.TAB_CHARS)
        Else
            Exit For
        End If
    Next i
    LeadingColumns = c
End Function
