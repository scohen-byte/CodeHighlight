Attribute VB_Name = "modGutter"
'==============================================================================
' modGutter - the line-number gutter.
'
' A SEPARATE text box, not numbers inside the code text. The block's text has to
' stay pure source: it is what the lexer reads and what gets copied out, and a
' "1 " on the front of every line would break both.
'
' The gutter sits INSIDE the dark block rather than beside it, which is what
' makes it look like an editor. That is done by giving the code text frame a
' left margin of padding + gutter width, and overlaying the numbers on the
' block's left edge - the same arrangement tools/lab.py renders.
'
' Alignment needs three things to hold, and all three are already true:
'   line spacing is set in EXACT points, so both shapes step identically
'   word wrap is off, so one paragraph is one visual line
'   paragraph spacing is zero on both
' If the two ever drift apart, pressing Stylize re-syncs them. That is the
' recovery path, and it means alignment need not be perfect first time.
'
' The shapes are deliberately NOT grouped: grouping makes editing the text
' awkward and complicates every other operation.
'==============================================================================
Option Explicit

Public Const TAG_GUTTER_OF As String = "CODEBLOCK_GUTTER_OF"

' The number the first NUMBERED line gets. Stored on the block, because it is a
' property of this snippet rather than of the deck: code split across three
' slides wants 1, then 24, then 51, and each block has to remember its own.
Public Const TAG_FIRST_LINE As String = "CODEBLOCK_FIRST_LINE"

Public Function FirstLine(ByVal shp As Shape) As Long
    Dim v As Long
    v = CLng(Val(shp.Tags(TAG_FIRST_LINE)))
    If v < 1 Then v = 1
    FirstLine = v
End Function

Public Sub SetFirstLine(ByVal shp As Shape, ByVal n As Long)
    If n < 1 Then n = 1
    shp.Tags.Add TAG_FIRST_LINE, CStr(n)
End Sub

Public Function FindGutter(ByVal shp As Shape) As Shape
    Dim sld As Slide, g As Shape, blockId As String

    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Function

    Set sld = OwningSlide(shp)
    ' Descends into groups: once Stylize has grouped the block with its parts,
    ' the gutter is no longer a top-level shape.
    For Each g In modBlock.AllShapes(sld)
        If g.Tags(TAG_GUTTER_OF) = blockId Then
            Set FindGutter = g
            Exit Function
        End If
    Next g
End Function

' A shape inside a group has the GROUP as its Parent, not the slide, so walk up
' until a Slide is reached.
Public Function OwningSlide(ByVal shp As Shape) As Slide
    Dim o As Object
    Set o = shp.Parent
    Do While TypeName(o) = "Shape"
        Set o = o.Parent
    Loop
    Set OwningSlide = o
End Function

Public Function HasGutter(ByVal shp As Shape) As Boolean
    HasGutter = Not (FindGutter(shp) Is Nothing)
End Function

' Creates the gutter if absent, removes it if present. Returns True when the
' gutter is on afterwards.
Public Function ToggleGutter(ByVal shp As Shape) As Boolean
    If HasGutter(shp) Then
        RemoveGutter shp
        ToggleGutter = False
    Else
        SyncGutter shp, True
        ToggleGutter = True
    End If
End Function

Public Sub RemoveGutter(ByVal shp As Shape)
    Dim g As Shape
    Set g = FindGutter(shp)
    If Not g Is Nothing Then g.Delete
    ' Hand the space back to the code.
    shp.TextFrame.MarginLeft = modSpec.SpecPad(modBlock.BlockFontSize(shp))
End Sub

' The number column, one entry per paragraph so the two shapes stay in step.
'
' Blank lines at the START or END are padding - people use them to give a block
' a little breathing room - so they get no number. Everything between the first
' and last non-blank line is numbered consecutively, INCLUDING blank lines in
' the middle, which are part of the code, and including the last line, which is
' the whole point.
Public Function NumberColumn(ByVal text As String, _
                             Optional ByVal startAt As Long = 1) As String
    Dim lines() As String, i As Long, firstReal As Long, lastReal As Long
    Dim n As Long, out As String

    If startAt < 1 Then startAt = 1
    n = startAt - 1
    lines = modBlock.SplitLines(text)

    firstReal = -1
    lastReal = -1
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            If firstReal < 0 Then firstReal = i
            lastReal = i
        End If
    Next i

    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then out = out & vbCr
        If firstReal >= 0 And i >= firstReal And i <= lastReal Then
            n = n + 1
            out = out & CStr(n)
        End If
    Next i

    NumberColumn = out
End Function

' Moves anything that is not a line number out of the gutter and onto the end of
' the code, where the user meant to type it.
Private Sub RescueStrayText(ByVal shp As Shape, ByVal g As Shape)
    Dim existing As String, lines() As String, i As Long, stray As String

    On Error Resume Next
    existing = g.TextFrame.TextRange.text
    On Error GoTo 0
    If Len(existing) = 0 Then Exit Sub

    lines = Split(Replace(Replace(existing, vbCrLf, vbCr), vbLf, vbCr), vbCr)
    For i = LBound(lines) To UBound(lines)
        ' A line number is digits and nothing else. Anything else was typed.
        If Len(Trim$(lines(i))) > 0 And Not IsNumeric(Trim$(lines(i))) Then
            If Len(stray) > 0 Then stray = stray & vbCr
            stray = stray & lines(i)
        End If
    Next i

    If Len(stray) = 0 Then Exit Sub
    shp.TextFrame.TextRange.text = shp.TextFrame.TextRange.text & vbCr & stray
End Sub

' Renumbers and re-aligns. Called on every Stylize, so the gutter follows the
' block's size, position and line count without the user thinking about it.
' create=False leaves a block without a gutter alone.
Public Sub SyncGutter(ByVal shp As Shape, Optional ByVal create As Boolean = False)
    Dim sld As Slide, g As Shape
    Dim size As Single, pad As Single, gutterW As Single, gap As Single
    Dim lineCount As Long, i As Long, numbers As String

    Set g = FindGutter(shp)
    If g Is Nothing And Not create Then Exit Sub

    Set sld = OwningSlide(shp)
    size = modBlock.BlockFontSize(shp)
    pad = modSpec.SpecPad(size)
    lineCount = modBlock.CountLines(shp.TextFrame.TextRange.text)
    ' Sized from the highest number that will appear, not from the line count.
    ' A twelve-line block starting at 98 needs three digits, and a gutter sized
    ' for two clips them - which reads as a numbering bug and is a width bug.
    gutterW = modSpec.SpecGutter(size, FirstLine(shp) + lineCount - 1)
    gap = Round(size * 0.45, 1)                 ' the gutter-to-code gap

    ' The code makes room for the numbers. With autofit on, widening this margin
    ' widens the block, so the gutter never overlaps the code.
    shp.TextFrame.MarginLeft = pad + gutterW

    If g Is Nothing Then
        Set g = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                      shp.Left, shp.Top, pad + gutterW, shp.Height)
        g.Tags.Add TAG_GUTTER_OF, shp.Tags(modBlock.TAG_ID)
        g.Line.Visible = msoFalse
        g.Fill.Visible = msoFalse
    End If

    ' RESCUE anything typed into the gutter before overwriting it.
    '
    ' The gutter is a transparent text box lying over the block's left edge, so
    ' it is entirely possible to click into it and type - the text then appears
    ' left of the code, uncoloured and unnumbered, because it is not in the code
    ' block at all. Overwriting it with the numbers would silently destroy it.
    RescueStrayText shp, g

    numbers = NumberColumn(shp.TextFrame.TextRange.text, FirstLine(shp))

    With g.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeNone
        .VerticalAnchor = msoAnchorTop
        .MarginLeft = 0
        .MarginRight = gap
        .MarginTop = pad
        .MarginBottom = pad
        .TextRange.text = numbers
        With .TextRange
            .Font.Name = THEME_FONT
            .Font.size = size
            .Font.Color.RGB = ThemeGutterColor()
            With .ParagraphFormat
                ' Right-aligned so the digits line up on their units column,
                ' the way an editor shows them.
                .Alignment = ppAlignRight
                .LineRuleWithin = msoFalse
                .SpaceWithin = modSpec.SpecLine(size)
                .LineRuleBefore = msoFalse
                .SpaceBefore = 0
                .LineRuleAfter = msoFalse
                .SpaceAfter = 0
            End With
        End With
    End With

    ' Size the gutter from the CONTENT, not from shp.Height.
    '
    ' Autofit recalculates the block's height lazily, so reading it here can
    ' return the height from BEFORE the last edit. A gutter sized from a stale
    ' height is too short, and PowerPoint then clips the final number - which
    ' looks exactly like an off-by-one in the numbering and is not one.
    g.Left = shp.Left
    g.Top = shp.Top
    g.Width = pad + gutterW
    g.Height = modSpec.SpecHeight(size, lineCount)
    If shp.Height > g.Height Then g.Height = shp.Height

    ' In FRONT of the block, or the block's opaque fill hides the numbers.
    ' Grouping and ungrouping reorders shapes, so this cannot be left to the
    ' order things happened to be created in.
    g.ZOrder msoBringToFront

    AlignFirstLine shp, g
End Sub

' Nudges the gutter so its first number sits on the SAME baseline as the first
' line of code.
'
' Both frames use the same font, size, exact line spacing and top margin, so in
' principle they should already agree. In practice they do not quite: the block
' has autofit on and the gutter does not, and PowerPoint lays the two frames out
' slightly differently as a result. Rather than model that, ask where the two
' first lines actually landed and close the gap.
Private Sub AlignFirstLine(ByVal shp As Shape, ByVal g As Shape)
    Dim codeTop As Single, numTop As Single, delta As Single

    On Error GoTo Done
    If Len(g.TextFrame.TextRange.text) = 0 Then Exit Sub
    If Len(shp.TextFrame.TextRange.text) = 0 Then Exit Sub

    codeTop = shp.TextFrame.TextRange.Characters(1, 1).BoundTop
    numTop = g.TextFrame.TextRange.Characters(1, 1).BoundTop
    delta = codeTop - numTop

    ' Only correct a real difference, and never a wild one - a bad measurement
    ' should leave the gutter where it was rather than throw it off the block.
    If Abs(delta) > 0.2 And Abs(delta) < modSpec.SpecLine(modBlock.BlockFontSize(shp)) Then
        g.Top = g.Top + delta
    End If
Done:
End Sub
