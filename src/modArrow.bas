Attribute VB_Name = "modArrow"
'==============================================================================
' modArrow - a block arrow in the left margin, pointing at one line.
'
' The other half of Emphasize, and a different instrument. Emphasize says
' "ignore everything else": it bands one line and fades the rest, which is right
' when the surroundings are noise. An arrow says "look here, and keep reading" -
' the code stays at full contrast and one line is singled out from outside it.
' A slide whose whole point is that the snippet is short and readable should not
' have most of it greyed out to draw the eye to line one.
'
' An arrow is an ORDINARY SHAPE tagged with its block and its line. Like a note
' it is placed rather than regenerated, so recolouring one by hand with
' PowerPoint's own tools survives every Stylize. Unlike a note it holds nothing
' the user typed, which is why Strip removes them and why clearing them all
' needs no confirmation - one click puts one back.
'
' Position comes from MEASURING the line, not from predicting it. modRender
' owns that, because the cover panels needed it first: exact line spacing puts
' the baseline where the spacing says and lets descenders hang below the
' computed box, so an arrow centred on the computed box sits slightly high.
'==============================================================================
Option Explicit

Public Const TAG_ARROW_OF   As String = "CODEBLOCK_ARROW_OF"
Public Const TAG_ARROW_LINE As String = "CODEBLOCK_ARROW_LINE"

Private Const GAP_PT    As Single = 8       ' arrow tip to the block's edge
Private Const MARGIN_PT As Single = 4       ' slide edge the arrow keeps clear
Private Const W_MIN     As Single = 14      ' below this it stops reading as an arrow

' A block at the default position leaves barely thirty points to its left, so
' the arrow is usually width-limited rather than drawn at its ideal proportions.
' It is therefore made TALL for its width - a stubby arrow still reads as an
' arrow, a thin one reads as a line.
Private Const H_RATIO   As Single = 0.72    ' of one line's height
Private Const W_RATIO   As Single = 1.9     ' of the arrow's own height
Private Const SHAFT     As Single = 0.45    ' shaft thickness, of the height
Private Const HEAD      As Single = 0.55    ' head length, of the width

'------------------------------------------------------------------------------
' Finding
'------------------------------------------------------------------------------

Public Function AllArrows(ByVal shp As Shape) As Collection
    Dim sld As Slide, s2 As Shape, blockId As String, c As Collection

    Set c = New Collection
    Set AllArrows = c
    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Function
    Set sld = modGutter.OwningSlide(shp)

    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(TAG_ARROW_OF) = blockId Then c.Add s2
    Next s2
End Function

Public Function ArrowCount(ByVal shp As Shape) As Long
    ArrowCount = AllArrows(shp).count
End Function

' The lowest line carrying an arrow, or 0. The renderer bolds it, the same way
' it bolds the newest emphasised line - whatever the slide is singling out is
' what should be heaviest.
Public Function LastArrowLine(ByVal shp As Shape) As Long
    Dim c As Collection, i As Long, v As Long
    Set c = AllArrows(shp)
    For i = 1 To c.count
        v = CLng(Val(c(i).Tags(TAG_ARROW_LINE)))
        If v > LastArrowLine Then LastArrowLine = v
    Next i
End Function

Public Function FindArrow(ByVal shp As Shape, ByVal lineNo As Long) As Shape
    Dim c As Collection, i As Long
    Set c = AllArrows(shp)
    For i = 1 To c.count
        If CLng(Val(c(i).Tags(TAG_ARROW_LINE))) = lineNo Then
            Set FindArrow = c(i)
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Adding and removing
'------------------------------------------------------------------------------

' Adds an arrow for one line, or returns the one already there. Geometry is left
' to PlaceArrows, which the caller reaches through StyleBlock.
Public Function AddArrow(ByVal shp As Shape, ByVal lineNo As Long) As Shape
    Dim sld As Slide, a As Shape

    Set a = FindArrow(shp, lineNo)
    If Not a Is Nothing Then
        Set AddArrow = a
        Exit Function
    End If

    Set sld = modGutter.OwningSlide(shp)
    Set a = sld.Shapes.AddShape(msoShapeRightArrow, shp.Left - 40, shp.Top, 30, 14)
    With a
        .fill.Solid
        .fill.ForeColor.RGB = modOptions.ArrowColor()
        .fill.Transparency = 0
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Tags.Add TAG_ARROW_OF, shp.Tags(modBlock.TAG_ID)
        .Tags.Add TAG_ARROW_LINE, CStr(lineNo)
    End With

    Set AddArrow = a
End Function

' True when there was one to remove, so the caller can report a toggle.
Public Function RemoveArrow(ByVal shp As Shape, ByVal lineNo As Long) As Boolean
    Dim a As Shape
    Set a = FindArrow(shp, lineNo)
    If a Is Nothing Then Exit Function
    a.Delete
    RemoveArrow = True
End Function

Public Sub ClearArrows(ByVal shp As Shape)
    Dim c As Collection, i As Long
    Set c = AllArrows(shp)
    For i = c.count To 1 Step -1
        c(i).Delete
    Next i
End Sub

' Repaints every arrow in the whole presentation.
'
' Deck-wide because arrows are uniform by design - see modOptions.ArrowColor -
' and because the alternative is worse: a walkthrough of twenty slides built
' before you settled on a colour would otherwise have to be recoloured slide by
' slide.
'
' Colour is applied on CREATION and here, never in PlaceArrows. Repainting on
' every Stylize would undo any arrow recoloured by hand with PowerPoint's own
' tools, which is the same bargain notes get.
Public Function RecolorAll(ByVal pres As Presentation, ByVal rgbColor As Long) As Long
    Dim sld As Slide, s2 As Shape, n As Long

    On Error Resume Next
    For Each sld In pres.Slides
        For Each s2 In modBlock.AllShapes(sld)
            If Len(s2.Tags(TAG_ARROW_OF)) > 0 Then
                s2.fill.Solid
                s2.fill.ForeColor.RGB = rgbColor
                n = n + 1
            End If
        Next s2
    Next sld
    RecolorAll = n
End Function

'------------------------------------------------------------------------------
' Placement
'------------------------------------------------------------------------------

' Sizes and positions every arrow from its line. Called after the block has
' reached its final geometry, so the arrow lands level with the line as
' PowerPoint actually laid it out.
Public Sub PlaceArrows(ByVal shp As Shape)
    Dim c As Collection, i As Long, ln As Long
    Dim size As Single, lineH As Single, pad As Single
    Dim h As Single, w As Single, room As Single
    Dim topPt As Single, heightPt As Single, midY As Single

    On Error GoTo Done
    Set c = AllArrows(shp)
    If c.count = 0 Then Exit Sub

    size = modBlock.BlockFontSize(shp)
    lineH = modSpec.SpecLine(size)
    pad = modSpec.SpecPad(size)

    ' Proportioned from the line, so an arrow beside 32pt code is not the same
    ' arrow as beside 14pt code.
    h = Round(lineH * H_RATIO, 1)
    w = Round(h * W_RATIO, 1)

    ' Whatever room there is to the left of the block. A block dragged near the
    ' left edge gets a shorter arrow rather than one hanging off the slide, and
    ' only a block with no margin at all gets none.
    room = shp.Left - MARGIN_PT - GAP_PT
    If w > room Then w = room

    For i = 1 To c.count
        ln = CLng(Val(c(i).Tags(TAG_ARROW_LINE)))
        If ln < 1 Then ln = 1

        ' Measured, with the computed line box as the fallback for a blank line
        ' or a failed measurement.
        If modRender.LineBounds(shp, ln, topPt, heightPt) Then
            midY = topPt + heightPt / 2
        Else
            midY = shp.Top + pad + (ln - 0.5) * lineH
        End If

        If w < W_MIN Then
            ' Nowhere to put it. Parked off-slide would be invisible and
            ' undraggable, so it goes at the margin at minimum size and overlaps
            ' - visibly wrong, which is better than silently absent.
            c(i).Width = W_MIN
            c(i).Left = MARGIN_PT
        Else
            c(i).Width = w
            c(i).Left = shp.Left - GAP_PT - w
        End If
        c(i).Height = h
        c(i).Top = midY - h / 2

        ' Set explicitly rather than left at PowerPoint's defaults, which are
        ' fractions of the CURRENT size: squeezing the width would otherwise
        ' shrink the head along with it and leave a thin dash.
        On Error Resume Next
        c(i).Adjustments(1) = SHAFT
        c(i).Adjustments(2) = HEAD
        On Error GoTo Done

        c(i).ZOrder msoBringToFront
    Next i
Done:
End Sub
