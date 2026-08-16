Attribute VB_Name = "modNote"
'==============================================================================
' modNote - notes attached to a line of code.
'
' A note is an ORDINARY TEXT SHAPE, tagged with the block it belongs to and the
' line it explains. It is never regenerated: the words in it are the user's, and
' the one thing this module must never do is retype them. Everything else about
' a note - where it sits, how wide it is, the leader line pointing at the code -
' is redrawn on every Stylize, the same way the gutter and the guides are.
'
' WHY THE POSITION IS TRACKED RATHER THAN RECOMPUTED.
'
' Notes need both things at once. They have to FOLLOW the code: change the font
' size and the line they point at moves, so the note must move with it. And they
' have to STAY PUT: two notes on adjacent lines cannot both sit centred on their
' own line, so the user drags one clear, and Stylize must not undo that.
'
' So each note stores the block's position and its own, as of the last time this
' module placed it. On the next pass the two deltas are compared:
'
'   the block moved by D and the note moved by D   -> the group was dragged
'   the block did not move and the note did        -> the user moved the note
'
' The difference is accumulated as the note's own offset from its anchor, and
' the note is then placed at anchor + offset. That is why CaptureDrags has to run
' BEFORE anything in the pipeline moves the block, and PlaceNotes after the block
' has reached its final geometry. Splitting them is not tidiness - measuring at
' either single moment gives the wrong answer for one of the two cases.
'
' Positions are stored rounded to whole points. A tag is a string, and CStr on a
' Single writes the LOCALE decimal separator, so "12.5" becomes "12,5" on a
' German machine and the comma-separated tag no longer parses. Whole points have
' no separator to get wrong, and half a point is invisible.
'==============================================================================
Option Explicit

Public Const TAG_NOTE_OF   As String = "CODEBLOCK_NOTE_OF"
Public Const TAG_NOTE_LINE As String = "CODEBLOCK_NOTE_LINE"
' The note's offset from its anchor, "dx,dy". Written ONLY when the user has
' moved the note - see CaptureDrags.
Public Const TAG_NOTE_OFF  As String = "CODEBLOCK_NOTE_OFF"
' What this module last did, as
' "blockLeft,blockTop,baseLeft,baseTop,placedLeft,placedTop".
'
' Base and placed are both needed, and they differ when the stacking pushed the
' note down. Comparing the note against PLACED says whether the user has moved
' it; measuring the new offset against BASE is what makes the offset mean "from
' where this note belongs". Storing only one of the two conflates a note that
' was stacked with a note that was dragged, and a stacked note then freezes and
' stops restacking when the note above it grows.
Public Const TAG_NOTE_SEEN As String = "CODEBLOCK_NOTE_SEEN"
' The thin connector from the note back to its line of code, and the invisible
' one-point shape at that line which gives the connector something to attach to.
Public Const TAG_LEADER_OF As String = "CODEBLOCK_LEADER_OF"
Public Const TAG_ANCHOR_OF As String = "CODEBLOCK_ANCHOR_OF"

Public Const NOTE_PLACEHOLDER As String = "What this line does"

' The size, colour and font new notes get live in modOptions, on the
' presentation, so the choice survives closing PowerPoint and travels with the
' deck.

Private Const GAP_PT    As Single = 24      ' block edge to note
Private Const MARGIN_PT As Single = 18      ' slide edge the note keeps clear
Private Const W_MAX     As Single = 260
Private Const W_MIN     As Single = 150
Private Const STACK_GAP As Single = 8       ' between two auto-placed notes
Private Const ANCHOR_PT As Single = 1       ' the invisible target on the line

' Rounding noise, not a drag. Positions are stored as whole points, so two
' roundings can differ by a point without anything having moved.
Private Const DRAG_MIN  As Single = 1.5

' CaptureDrags must run EXACTLY ONCE per placement pass, and it must run while
' the block is still where the user left it.
'
' Once is not a style point. The capture reads "the block moved and the note did
' not" as a drag, and that is true only of a drag - Resize moves the block on
' purpose, with the notes ungrouped, so a second capture after it would record
' the resize as though the user had dragged every note backwards.
'
' Commands that move the block themselves therefore capture first and then call
' the ordinary pipeline, whose own capture becomes a no-op. PlaceNotes closes
' the pass.
Private mCaptured As Boolean

'------------------------------------------------------------------------------
' Finding notes
'------------------------------------------------------------------------------

' Every note belonging to a block, in LINE order, 1-based. Returns the count.
'
' Line order is what makes the stacking deterministic: notes are placed top
' down, and each one that has not been dragged is kept below the one before it.
Public Function NoteArray(ByVal shp As Shape, ByRef arr() As Shape) As Long
    Dim sld As Slide, s2 As Shape, blockId As String
    Dim n As Long, i As Long, j As Long, tmp As Shape

    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Function
    Set sld = modGutter.OwningSlide(shp)

    ReDim arr(1 To 1)
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(TAG_NOTE_OF) = blockId Then
            n = n + 1
            ReDim Preserve arr(1 To n)
            Set arr(n) = s2
        End If
    Next s2

    For i = 2 To n
        Set tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If NoteLine(arr(j)) <= NoteLine(tmp) Then Exit Do
            Set arr(j + 1) = arr(j)
            j = j - 1
        Loop
        Set arr(j + 1) = tmp
    Next i

    NoteArray = n
End Function

Public Function NoteLine(ByVal note As Shape) As Long
    NoteLine = CLng(Val(note.Tags(TAG_NOTE_LINE)))
End Function

Public Function NoteCount(ByVal shp As Shape) As Long
    Dim arr() As Shape
    NoteCount = NoteArray(shp, arr)
End Function

Public Function FindNote(ByVal shp As Shape, ByVal lineNo As Long) As Shape
    Dim arr() As Shape, n As Long, i As Long
    n = NoteArray(shp, arr)
    For i = 1 To n
        If NoteLine(arr(i)) = lineNo Then
            Set FindNote = arr(i)
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Creating and removing
'------------------------------------------------------------------------------

' Adds a note for one line, carrying placeholder text. Position is left to
' PlaceNotes, which the caller reaches through StyleBlock.
Public Function AddNote(ByVal shp As Shape, ByVal lineNo As Long) As Shape
    Dim sld As Slide, r As Shape
    Dim nSize As Single, w As Single, pad As Single, fill As Long, fnt As String

    Set sld = modGutter.OwningSlide(shp)
    nSize = EffectiveNoteSize(shp)
    fill = modOptions.NoteColor()
    fnt = modOptions.NoteFont()
    pad = modSpec.SpecPad(nSize)
    w = NoteWidth(shp)

    Set r = sld.Shapes.AddShape(msoShapeRoundedRectangle, _
                                shp.Left + shp.Width + GAP_PT, shp.Top, w, nSize * 3)
    With r
        .Fill.Solid
        .Fill.ForeColor.RGB = fill
        .Fill.Transparency = 0
        .Shadow.Visible = msoFalse
        .Tags.Add TAG_NOTE_OF, shp.Tags(modBlock.TAG_ID)
        .Tags.Add TAG_NOTE_LINE, CStr(lineNo)
    End With

    With r.TextFrame
        ' Wrap ON, unlike the code block: a note is prose, and its width is
        ' fixed so that it fits beside the block. Autofit then grows the height.
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = pad
        .MarginRight = pad
        .MarginTop = Round(pad * 0.6, 1)
        .MarginBottom = Round(pad * 0.6, 1)
        .TextRange.text = NOTE_PLACEHOLDER
        With .TextRange
            ' The font is deliberately NOT the code font. Left unset it
            ' inherits the deck's own body font, so a note looks like the rest
            ' of the presentation rather than like this add-in - which is the
            ' default, and what Note font calls "Deck default".
            If Len(fnt) > 0 Then .Font.Name = fnt
            .Font.size = nSize
            .Font.Color.RGB = ThemeTextOn(fill)
            .ParagraphFormat.Alignment = ppAlignLeft
        End With
        ' Autofit last, so the first height is measured from the real text and
        ' the real margins rather than from the placeholder box.
        .AutoSize = ppAutoSizeShapeToFitText
    End With

    r.Width = w
    r.Adjustments(1) = modSpec.SpecCornerAdjust(nSize, ShorterSide(r))
    ApplyEdge r, fill

    Set AddNote = r
End Function

' A light note on a light slide has no edge of its own and reads as words
' floating in space. A dark one does not need the help and looks busier for it.
Private Sub ApplyEdge(ByVal note As Shape, ByVal fill As Long)
    If ThemeNeedsEdge(fill) Then
        note.Line.Visible = msoTrue
        note.Line.ForeColor.RGB = ThemeEdgeFor(fill)
        note.Line.Weight = 1
    Else
        note.Line.Visible = msoFalse
    End If
End Sub

Public Sub RemoveNote(ByVal shp As Shape, ByVal lineNo As Long)
    Dim note As Shape
    Set note = FindNote(shp, lineNo)
    If Not note Is Nothing Then note.Delete
End Sub

' Deletes every note on the block. Destructive of typed text, so the caller asks
' first - this is the one thing in the add-in that can lose words.
Public Sub ClearNotes(ByVal shp As Shape)
    Dim arr() As Shape, n As Long, i As Long
    ClearLeaders shp
    n = NoteArray(shp, arr)
    For i = n To 1 Step -1
        arr(i).Delete
    Next i
End Sub

Public Sub ClearLeaders(ByVal shp As Shape)
    Dim sld As Slide, blockId As String, doomed As Collection, s2 As Shape, i As Long

    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Sub
    Set sld = modGutter.OwningSlide(shp)

    ' Connectors BEFORE anchors. Deleting a shape a connector is attached to
    ' leaves the connector behind with a dangling end, so the order is not
    ' arbitrary - and AllShapes is a snapshot, so both are collected first and
    ' deleted afterwards either way.
    Set doomed = New Collection
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(TAG_LEADER_OF) = blockId Then doomed.Add s2
    Next s2
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(TAG_ANCHOR_OF) = blockId Then doomed.Add s2
    Next s2
    For i = doomed.count To 1 Step -1
        doomed(i).Delete
    Next i
End Sub

'------------------------------------------------------------------------------
' Placement
'------------------------------------------------------------------------------

' Works out which notes the user has dragged since the last pass, and records
' that as each note's offset from where this module would otherwise put it.
' MUST run before the block moves - see the module header.
'
' The offset is REPLACED, not accumulated - one subtraction against the stored
' base gives the whole answer, where accumulating deltas would make every
' rounding error permanent. And it is written ONLY when the note has actually
' moved, so a note nobody has touched keeps no offset at all and goes on taking
' part in the stacking.
Public Sub CaptureDrags(ByVal shp As Shape)
    Dim arr() As Shape, n As Long, i As Long, v() As String
    Dim bdx As Single, bdy As Single, mx As Single, my As Single

    If mCaptured Then Exit Sub
    mCaptured = True

    On Error GoTo Done
    n = NoteArray(shp, arr)
    For i = 1 To n
        v = Split(arr(i).Tags(TAG_NOTE_SEEN), ",")
        If UBound(v) = 5 Then
            ' Whatever the block did, the note did too if they moved together.
            ' Subtracting the block's own movement is what separates "the group
            ' was dragged" from "this note was dragged".
            bdx = shp.Left - Val(v(0))
            bdy = shp.Top - Val(v(1))
            If Abs(bdx) < DRAG_MIN Then bdx = 0
            If Abs(bdy) < DRAG_MIN Then bdy = 0

            ' How far the note is from where this module last put it. Below the
            ' threshold it has not moved, and the difference is the half point
            ' lost to storing positions as whole numbers.
            mx = (arr(i).Left - bdx) - Val(v(4))
            my = (arr(i).Top - bdy) - Val(v(5))

            If Abs(mx) >= DRAG_MIN Or Abs(my) >= DRAG_MIN Then
                arr(i).Tags.Add TAG_NOTE_OFF, _
                    PtStr((arr(i).Left - bdx) - Val(v(2))) & "," & _
                    PtStr((arr(i).Top - bdy) - Val(v(3)))
            End If
        End If
    Next i
Done:
End Sub

' Puts every note where it belongs and redraws its leader line.
'
' A note is placed at BASE + OFFSET, where the base is worked out afresh from
' the block's geometry and the offset is whatever the note has drifted from it.
' The base is rounded to whole points before use, so that it matches the copy
' stored in the tag exactly - otherwise half a point of rounding is added to the
' offset on every Stylize, and a note walks slowly across the slide.
Public Sub PlaceNotes(ByVal shp As Shape)
    Dim sld As Slide, arr() As Shape, n As Long, i As Long
    Dim size As Single, lineH As Single, pad As Single
    Dim ax As Single, ay As Single, ln As Long
    Dim baseX As Single, baseY As Single, x As Single, y As Single
    Dim floorY As Single, offX As Single, offY As Single
    Dim free As Boolean, stacked As Boolean

    On Error GoTo Done
    ClearLeaders shp
    n = NoteArray(shp, arr)
    If n = 0 Then GoTo Done

    Set sld = modGutter.OwningSlide(shp)
    size = modBlock.BlockFontSize(shp)
    lineH = modSpec.SpecLine(size)
    pad = modSpec.SpecPad(size)
    floorY = -10000

    For i = 1 To n
        ln = NoteLine(arr(i))
        If ln < 1 Then ln = 1

        ' The anchor is the block's right edge, level with the middle of the
        ' line. Everything else is measured from there.
        ax = shp.Left + shp.Width
        ay = shp.Top + pad + (ln - 0.5) * lineH

        If ax + GAP_PT + arr(i).Width <= modSpec.SLIDE_W - MARGIN_PT Then
            baseX = ax + GAP_PT
            baseY = ay - arr(i).Height / 2
        ElseIf shp.Left - GAP_PT - arr(i).Width >= MARGIN_PT Then
            ' No room on the right. The left margin is the next best place: it
            ' still lets the note sit level with its line.
            baseX = shp.Left - GAP_PT - arr(i).Width
            baseY = ay - arr(i).Height / 2
        Else
            ' A block this wide leaves no margin at all, so the note goes below
            ' it, right-aligned, and the leader carries the whole burden of
            ' saying which line is meant.
            baseX = ax - arr(i).Width
            baseY = shp.Top + shp.Height + GAP_PT
        End If

        baseX = CSng(CLng(baseX))
        baseY = CSng(CLng(baseY))

        free = Not ParseOffset(arr(i), offX, offY)
        x = baseX + offX
        y = baseY + offY
        stacked = False

        If free Then
            ' Two notes on nearby lines would land on top of each other. The
            ' lower one gives way, which keeps both readable without the user
            ' having to drag anything.
            If y < floorY Then
                y = floorY
                stacked = True
            End If

            ' Only a note that is still where this module put it gets pulled
            ' back onto the slide. Once it carries an offset from a drag, its
            ' position is the answer and nudging it would fight the user.
            If x < MARGIN_PT Then x = MARGIN_PT
            If x + arr(i).Width > modSpec.SLIDE_W - MARGIN_PT Then
                x = modSpec.SLIDE_W - MARGIN_PT - arr(i).Width
            End If
            If y < MARGIN_PT Then y = MARGIN_PT

            ' The bottom edge does NOT win against the stack. Pulling the last
            ' note up to fit puts it on top of the one above, which reads as a
            ' rendering fault; letting it hang over the edge reads as a slide
            ' with too much on it, which is what it is. DoNote says so.
            If Not stacked Then
                If y + arr(i).Height > modSpec.SLIDE_H - MARGIN_PT Then
                    y = modSpec.SLIDE_H - MARGIN_PT - arr(i).Height
                End If
            End If
        End If

        arr(i).Left = x
        arr(i).Top = y
        floorY = y + arr(i).Height + STACK_GAP

        ' Base AND final position. CaptureDrags needs both to tell a note the
        ' stacking moved from one the user moved.
        arr(i).Tags.Add TAG_NOTE_SEEN, _
                        PtStr(shp.Left) & "," & PtStr(shp.Top) & "," & _
                        PtStr(baseX) & "," & PtStr(baseY) & "," & _
                        PtStr(x) & "," & PtStr(y)

        AddLeader sld, shp, arr(i), ax, ay
    Next i
Done:
    ' Closes the pass, on the error path too - a pass left open would make the
    ' NEXT Stylize skip its capture and quietly lose a drag.
    mCaptured = False
End Sub

' True when the note would sit off the bottom of the slide, which is the one
' placement failure the user has to be told about - a wide block leaves room for
' only two or three notes below it, and after that they run off the edge.
Public Function NoteOffSlide(ByVal note As Shape) As Boolean
    NoteOffSlide = (note.Top + note.Height > modSpec.SLIDE_H) Or (note.Top < 0)
End Function

' The leader is a real CONNECTOR, not a drawn line.
'
' A drawn line has fixed endpoints, so dragging a note left it pointing at
' nothing until the next Stylize. A connector attached to two shapes is rerouted
' by PowerPoint itself, live, as either end moves - which is the whole of what
' was wanted, and it is native rather than something to reimplement.
'
' Connecting it needs something AT THE LINE to connect to. A rectangle offers
' connection sites at its edge midpoints and corners, so connecting to the block
' would anchor the leader to the middle of the block's right edge rather than to
' line 7. So each leader gets a one-point invisible anchor placed at its line,
' and the connector runs anchor to note.
'
' RerouteConnections then picks the closest pair of sites on its own, which is
' exactly what the hand-rolled nearest-edge-point search used to do.
Private Sub AddLeader(ByVal sld As Slide, ByVal shp As Shape, ByVal note As Shape, _
                      ByVal ax As Single, ByVal ay As Single)
    Dim anchor As Shape, lin As Shape, blockId As String

    On Error GoTo Done
    blockId = shp.Tags(modBlock.TAG_ID)

    Set anchor = sld.Shapes.AddShape(msoShapeRectangle, ax, ay - ANCHOR_PT / 2, _
                                     ANCHOR_PT, ANCHOR_PT)
    With anchor
        .fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Tags.Add TAG_ANCHOR_OF, blockId
    End With

    Set lin = sld.Shapes.AddConnector(msoConnectorStraight, ax, ay, ax + 10, ay)
    With lin
        .ConnectorFormat.BeginConnect anchor, 1
        .ConnectorFormat.EndConnect note, 1
        ' Only now, with both ends attached, can PowerPoint choose sites.
        .RerouteConnections
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = ThemeLeaderColor()
        .Line.Weight = 1.25
        .Tags.Add TAG_LEADER_OF, blockId
        ' Behind everything: the leader starts ON the block's edge, so it has
        ' nothing to cover, and in front it would draw across the note's corner.
        .ZOrder msoSendToBack
    End With
Done:
End Sub

'------------------------------------------------------------------------------

Public Function NoteFontSize(ByVal blockSize As Single) As Single
    Dim s As Single
    ' Smaller than the code, because a note is an aside. Floored at the teaching
    ' size, because an aside nobody can read is worse than no aside.
    s = Round(blockSize * 0.8, 0)
    If s < modSpec.MIN_TEACHING_SIZE Then s = modSpec.MIN_TEACHING_SIZE
    ' Capped at 28 rather than 20. The old cap meant Auto returned the same
    ' twenty points for every block from 25pt upward, which is not tracking the
    ' block at all - and twenty points of proportional prose beside 32pt code
    ' reads as a caption rather than as a note.
    If s > 28 Then s = 28
    NoteFontSize = s
End Function

'------------------------------------------------------------------------------
' The size, colour and font new notes get. The choices themselves live in
' modOptions, on the presentation; what belongs here is applying them.
'------------------------------------------------------------------------------

' The size a note on this block should be: the pinned choice, or derived.
Public Function EffectiveNoteSize(ByVal shp As Shape) As Single
    Dim pinned As Long
    pinned = modOptions.NoteSize()
    If pinned >= 1 Then
        EffectiveNoteSize = CSng(pinned)
    Else
        EffectiveNoteSize = NoteFontSize(modBlock.BlockFontSize(shp))
    End If
End Function

' Repaints and resizes every note on a block. The text is untouched - only the
' size, the font, the fill and the text colour, which is derived from the fill
' so a light preset does not end up with near-white words on it.
'
' Height follows the text, so the caller has to place the notes again afterwards.
' ONE PROPERTY AT A TIME.
'
' There used to be a single RestyleNotes that reapplied size, colour and font
' together from the deck defaults. That is what forced every note on a slide to
' look the same: changing the colour also reset the size and the font, so no
' note could differ from any other in any respect. Each of these touches only
' what it is named after, and everything else about the note survives.
'
' Colour is stored nowhere but on the shape, which is what makes per-note colour
' work at all - Stylize places notes and never repaints them.
Public Sub ApplyFill(ByVal notes As Collection, ByVal fill As Long)
    Dim i As Long

    On Error Resume Next
    For i = 1 To notes.count
        With notes(i)
            .fill.Solid
            .fill.ForeColor.RGB = fill
            .fill.Transparency = 0
            ' The text colour is not a separate choice. It is whichever of light
            ' or dark reads better on this fill, so it has to move with it.
            .TextFrame.TextRange.Font.Color.RGB = ThemeTextOn(fill)
        End With
        ApplyEdge notes(i), fill
    Next i
End Sub

Public Sub ApplyFontSize(ByVal notes As Collection, ByVal pts As Single)
    Dim i As Long, pad As Single, w As Single

    On Error Resume Next
    pad = modSpec.SpecPad(pts)
    For i = 1 To notes.count
        ' Autofit reflows the height for the new size; the width is ours to keep.
        w = notes(i).Width
        With notes(i).TextFrame
            .MarginLeft = pad
            .MarginRight = pad
            .MarginTop = Round(pad * 0.6, 1)
            .MarginBottom = Round(pad * 0.6, 1)
            .TextRange.Font.size = pts
        End With
        notes(i).Width = w
        notes(i).Adjustments(1) = modSpec.SpecCornerAdjust(pts, ShorterSide(notes(i)))
    Next i
End Sub

' Empty means the deck's own body font. Setting Font.Name to "" is not the same
' thing - it is an error - so there is no way to put a note BACK to the deck
' font once it has been given a named one, short of retyping it. Choosing the
' deck default therefore leaves the selected notes alone and only affects the
' ones made afterwards.
Public Sub ApplyFontName(ByVal notes As Collection, ByVal fontName As String)
    Dim i As Long
    If Len(fontName) = 0 Then Exit Sub

    On Error Resume Next
    For i = 1 To notes.count
        notes(i).TextFrame.TextRange.Font.Name = fontName
    Next i
End Sub

Public Function IsNote(ByVal shp As Shape) As Boolean
    On Error Resume Next
    IsNote = (Len(shp.Tags(TAG_NOTE_OF)) > 0)
End Function

' The block a note belongs to, so a command that starts from a selected note can
' still run the pipeline that places it.
Public Function BlockOfNote(ByVal note As Shape) As Shape
    Dim sld As Slide, s2 As Shape, blockId As String

    blockId = note.Tags(TAG_NOTE_OF)
    If Len(blockId) = 0 Then Exit Function
    Set sld = modGutter.OwningSlide(note)

    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modBlock.TAG_BLOCK) = "1" Then
            If s2.Tags(modBlock.TAG_ID) = blockId Then
                Set BlockOfNote = s2
                Exit Function
            End If
        End If
    Next s2
End Function

' Every note on a block, as a Collection - the shape the styling subs above take.
Public Function AllNotes(ByVal shp As Shape) As Collection
    Dim arr() As Shape, n As Long, i As Long, c As Collection
    Set c = New Collection
    n = NoteArray(shp, arr)
    For i = 1 To n
        c.Add arr(i)
    Next i
    Set AllNotes = c
End Function

' How wide a new note can be: the room to the right of the block, then the room
' to its left, and failing both the full slide width for a note placed below.
Private Function NoteWidth(ByVal shp As Shape) As Single
    Dim avail As Single

    avail = modSpec.SLIDE_W - MARGIN_PT - (shp.Left + shp.Width + GAP_PT)
    If avail < W_MIN Then avail = shp.Left - MARGIN_PT - GAP_PT
    If avail < W_MIN Then avail = modSpec.SLIDE_W - 2 * MARGIN_PT
    If avail > W_MAX Then avail = W_MAX
    NoteWidth = avail
End Function

Private Function ShorterSide(ByVal shp As Shape) As Single
    ShorterSide = shp.Height
    If shp.Width < ShorterSide Then ShorterSide = shp.Width
End Function

' True when the note carries an offset, i.e. the user has moved it.
Private Function ParseOffset(ByVal note As Shape, ByRef dx As Single, ByRef dy As Single) As Boolean
    Dim v() As String
    dx = 0
    dy = 0
    v = Split(note.Tags(TAG_NOTE_OFF), ",")
    If UBound(v) <> 1 Then Exit Function
    dx = CSng(Val(v(0)))
    dy = CSng(Val(v(1)))
    ParseOffset = True
End Function

' Whole points, so the string has no decimal separator to be mangled by a
' locale. See the module header.
Private Function PtStr(ByVal v As Single) As String
    PtStr = CStr(CLng(v))
End Function
