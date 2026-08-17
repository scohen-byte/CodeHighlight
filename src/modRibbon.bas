Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon - ribbon callbacks, and the commands behind them.
'
' The callbacks are deliberately thin wrappers over Do* subs that take no
' arguments. IRibbonControl cannot be constructed from a test script, so a
' command with its logic inside the callback cannot be driven from anywhere but
' a mouse click. Splitting them means the whole command path is testable over
' COM - see modSelfTest.RibbonSliceTest.
'
' Ribbon callbacks must match the signatures Office expects exactly. A wrong
' signature fails at CLICK time with "Cannot run the macro", which looks like a
' packaging problem and is not.
'==============================================================================
Option Explicit

Private mRibbon    As IRibbonUI
Private mLangIndex As Long

' Warnings go through Warn so a test run cannot be blocked by a modal dialog.
' A MsgBox waiting for OK inside an automated run looks exactly like a hang.
Private mQuiet       As Boolean
Private mLastWarning As String

Public Const ADDIN_NAME As String = "CodeHighlight"

'------------------------------------------------------------------------------
' Load
'------------------------------------------------------------------------------

' Cached so the dropdown and the toggle can be re-read when the selection
' changes. PowerPoint gives an add-in no selection-change event, so for now the
' dropdown shows the last explicit choice rather than following the selection.
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    mLangIndex = 0
End Sub

Public Sub RefreshRibbon()
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

'------------------------------------------------------------------------------
' Commands. These are the real work, and they are callable from a script.
'------------------------------------------------------------------------------

Public Sub DoNewBlock()
    Dim sld As Slide, shp As Shape, lang As LangDef

    Set sld = ActiveSlide()
    If sld Is Nothing Then
        Warn "Open a slide in Normal view first."
        Exit Sub
    End If

    lang = modLangRegistry.LangAt(mLangIndex)
    Set shp = modBlock.CreateBlock(sld, PlaceholderFor(lang), modSpec.BASE_SIZE, lang.id)
    modRender.ApplyHighlight shp, lang.id

    ' Select the placeholder TEXT, not just the shape, so the first keystroke
    ' replaces it and there is nothing to delete by hand.
    On Error Resume Next
    shp.Select
    shp.TextFrame.TextRange.Select
    On Error GoTo 0
End Sub

Public Sub DoStylize()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, langId As String
    Dim before As String, after As String, size As Single

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    ' Out of the group before anything is redrawn, back into it at the end.
    modBlock.UngroupParts shp

    ' Repair autocorrect damage first, so the lexer sees code rather than
    ' typography. Rewriting the text drops run formatting, so the block
    ' formatting is reapplied before colouring.
    size = modBlock.BlockFontSize(shp)
    before = shp.TextFrame.TextRange.text
    after = modBlock.NormalizeCodeText(before)
    If after <> before Then
        shp.TextFrame.TextRange.text = after
        modBlock.FormatBlockText shp, size
    End If

    langId = modBlock.BlockLangId(shp, CurrentLangId())
    ' Adopts an untagged block - a block pasted from a colleague's deck starts
    ' working the first time it is highlighted.
    modBlock.EnsureTags shp, langId

    StyleBlock shp, langId
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoStylize failed: " & Err.Description
End Sub

' The whole styling pipeline for ONE block, with no reference to the selection.
'
' Separated out so it can be run on a block the user is not looking at - which
' is what Step through needs, since it styles blocks on slides it has just
' created.
Public Sub StyleBlock(ByVal shp As Shape, Optional ByVal langId As String = "")
    If Len(langId) = 0 Then langId = modBlock.BlockLangId(shp, CurrentLangId())
    ' FIRST, before anything moves: this is the only moment at which a note that
    ' the user dragged can be told apart from one that came along with the block.
    ' modNote's header explains why.
    modNote.CaptureDrags shp
    modBlock.UngroupParts shp
    ' Prompts live in the text now, so they are settled BEFORE anything is
    ' coloured - and assigning to TextRange.text drops every run, so the block
    ' formatting has to go back on when it changed.
    If modOutput.SyncPrompts(shp, langId) Then
        modBlock.FormatBlockText shp, modBlock.BlockFontSize(shp)
    End If
    modRender.ApplyHighlight shp, langId
    modBlock.ResizeToContent shp
    ' Gutter before guides: it changes the left margin, and the guides are
    ' placed from that margin.
    modGutter.SyncGutter shp
    modGuides.DrawGuides shp
    ' Covers next: they have to sit above everything to hide anything.
    modRender.DrawCovers shp
    ' Notes and arrows after the block has reached its final size and position,
    ' since both are placed from its edges.
    ' Output notes are dressed BEFORE they are placed. Dressing changes a note's
    ' width and height, and PlaceNotes records where it put each note so that
    ' the next pass can tell a drag from a stack - so resizing afterwards looks
    ' exactly like the user having dragged every output note, and they pile up.
    ' The wrap decision only needs the room beside the BLOCK, which is known
    ' without placing anything.
    modNote.SyncOutputNotes shp
    modNote.PlaceNotes shp
    modArrow.PlaceArrows shp
    ' Back into a group, so the block and its parts drag as one.
    modBlock.GroupParts shp
End Sub

Public Sub DoToggleGutter()
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    ' Style first. Toggling numbers on a block that was never styled would
    ' otherwise add them to uncoloured code, and the numbers would be laid out
    ' against geometry that has not been settled.
    modBlock.UngroupParts shp
    modGutter.ToggleGutter shp
    StyleBlock shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoToggleGutter failed: " & Err.Description
End Sub

' Sets the number the first numbered line gets, for code split across slides.
Public Sub DoFirstLine(ByVal n As Long)
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    modGutter.SetFirstLine shp, n
    ' Numbers the user cannot see are not much use, so turning this on turns the
    ' gutter on. Anyone setting a start number has said what they want.
    modBlock.UngroupParts shp
    If Not modGutter.HasGutter(shp) Then modGutter.SyncGutter shp, True
    StyleBlock shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoFirstLine failed: " & Err.Description
End Sub

Public Sub DoToggleGuides()
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    ' Style first, for the same reason as the numbers: guides are placed from
    ' the settled text geometry, so they need the block rendered before they can
    ' land in the right place.
    modBlock.UngroupParts shp
    modGuides.SetGuidesEnabled shp, Not modGuides.GuidesEnabled(shp)
    StyleBlock shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoToggleGuides failed: " & Err.Description
End Sub

' Emphasise the lines covered by the text selection.
'
' Select some text inside the block and press this. With the SHAPE selected
' rather than text, it clears the emphasis instead - so the same button both
' sets and removes, which is what you want when stepping through code slide by
' slide.
Public Sub DoEmphasize()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sel As Selection
    Dim txt As String, a As Long, b As Long, i As Long, list As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText Then
        If sel.TextRange.length > 0 Then
            txt = shp.TextFrame.TextRange.text
            a = modBlock.LineOfChar(txt, sel.TextRange.Start)
            b = modBlock.LineOfChar(txt, sel.TextRange.Start + sel.TextRange.length - 1)
            For i = a To b
                If Len(list) > 0 Then list = list & ","
                list = list & CStr(i)
            Next i
        End If
    End If

    modBlock.SetEmphasis shp, list
    modBlock.UngroupParts shp
    modRender.ApplyHighlight shp, modBlock.BlockLangId(shp, CurrentLangId())
    modBlock.GroupParts shp
    Reselect shp
    ' Clearing is silent. A dialog to dismiss on every step of a walkthrough is
    ' worse than no feedback at all - the block itself is the feedback.
    Exit Sub
Failed:
    Warn "DoEmphasize failed: " & Err.Description
End Sub

' Puts the block's code on the clipboard as text.
Public Sub DoCopyCode()
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    ' Copying the TextRange rather than the shape: pasting into an editor then
    ' gives the source, not a picture of a rounded rectangle.
    '
    ' A transcript's OUTPUT lines are dropped, for the same reason: what you
    ' paste should run. Doing that needs a scratch shape, because the clipboard
    ' takes a range and the range would bring the output with it.
    If modOutput.IsTranscript(shp) Then
        CopyTextViaScratch shp, _
            modOutput.CodeOnly(shp.TextFrame.TextRange.text, _
                               modOutput.GetOutputLines(shp), _
                               modBlock.BlockLangId(shp, CurrentLangId()))
    Else
        shp.TextFrame.TextRange.Copy
    End If
    Exit Sub
Failed:
    Warn "DoCopyCode failed: " & Err.Description
End Sub

' Puts arbitrary text on the clipboard by way of a shape that exists just long
' enough to be copied. VBA has no clipboard object without a reference to the
' Forms library, and this add-in adds no references - a missing one is a load
' failure on someone else's machine.
Private Sub CopyTextViaScratch(ByVal shp As Shape, ByVal text As String)
    Dim sld As Slide, tmp As Shape

    On Error GoTo Failed
    Set sld = modGutter.OwningSlide(shp)
    ' Off-slide, so it cannot flash into view on the way past.
    Set tmp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, -2000, -2000, 400, 100)
    tmp.TextFrame.TextRange.text = text
    tmp.TextFrame.TextRange.Copy
    tmp.Delete
    Exit Sub
Failed:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    shp.TextFrame.TextRange.Copy
End Sub

' Returns a block to plain, uncoloured text: no syntax colours, no emphasis, no
' line numbers, no guides. The block itself stays - same font, same dark fill -
' so this is an undo for the rendering, not for the block.
'
' The tags stay too, so pressing Stylize brings it all straight back.
Public Sub DoStrip()
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    modBlock.UngroupParts shp
    modGutter.RemoveGutter shp
    modGuides.SetGuidesEnabled shp, False
    modGuides.DrawGuides shp
    modBlock.SetEmphasis shp, ""

    shp.TextFrame.TextRange.Font.Color.RGB = ThemeColor(tkDefault)
    modRender.ClearBands shp
    modBlock.SetHidden shp, ""
    modRender.ClearCovers shp
    ' Arrows go, notes stay. An arrow holds nothing the user typed and one
    ' click puts one back, so it belongs with the emphasis it replaces.
    modArrow.ClearArrows shp
    ' Output marking goes too: it is a rendering choice about lines, like the
    ' emphasis and the covers, and Stylize brings it all back from the tag.
    modOutput.SetOutputLines shp, ""
    modOutput.SetTranscript shp, False

    modBlock.ResizeToContent shp
    ' Notes SURVIVE a strip. Everything else this removes can be rebuilt by
    ' pressing Stylize, and the words in a note cannot - so they are content,
    ' not styling. They are re-placed rather than left pointing at where the
    ' block used to be.
    modNote.PlaceNotes shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoStrip failed: " & Err.Description
End Sub

' Builds a walkthrough: one new slide per line of code, each emphasising the
' next step. The slide you start from is left with nothing emphasised, so it
' still shows the code whole before the walk begins.
'
' cumulative = False spotlights one line at a time, for tracing execution.
' cumulative = True grows the emphasis downward, for building code up.
' Hides the lines covered by the text selection, behind a panel with a question
' mark. With the shape selected rather than text, it reveals everything again.
Public Sub DoHide()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sel As Selection
    Dim txt As String, a As Long, b As Long, i As Long, list As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText Then
        If sel.TextRange.length > 0 Then
            txt = shp.TextFrame.TextRange.text
            a = modBlock.LineOfChar(txt, sel.TextRange.Start)
            b = modBlock.LineOfChar(txt, sel.TextRange.Start + sel.TextRange.length - 1)
            For i = a To b
                If Len(list) > 0 Then list = list & ","
                list = list & CStr(i)
            Next i
        End If
    End If

    modBlock.SetHidden shp, list
    modBlock.UngroupParts shp
    StyleBlock shp
    Reselect shp
    Exit Sub
Failed:
    Warn "DoHide failed: " & Err.Description
End Sub

' Attaches a note to the line the cursor is on, or opens the one already there.
'
' The cursor is enough - no selection needed - because a note is about one line,
' and asking someone to select the line first would be a step for nothing.
'
' With the BLOCK selected rather than text there is no cursor, so the line is
' taken from the emphasis. On a walkthrough slide that is exactly the line being
' talked about, which makes annotating a generated deck one click per slide with
' nothing to aim at. Build up emphasises a range, and the LAST line of it is the
' one the slide just arrived at.
'
' Only with no cursor and no emphasis does the gesture mean the other thing you
' might want: clear them. That one asks first. Every other clear in this add-in
' is silent, because everything else can be rebuilt by pressing Stylize, and
' typed words cannot.
Public Sub DoNote()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sel As Selection
    Dim ln As Long, note As Shape, n As Long

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText Then
        ln = modBlock.LineOfChar(shp.TextFrame.TextRange.text, sel.TextRange.Start)
    End If
    If ln < 1 Then ln = modBlock.LastEmphasisedLine(modBlock.GetEmphasis(shp))

    If ln < 1 Then
        n = modNote.NoteCount(shp)
        If n = 0 Then
            Warn "Click into the block, put the cursor on the line you want to " & _
                 "explain, and press Note. On a walkthrough slide the emphasised " & _
                 "line is used, so the block itself is enough."
            Exit Sub
        End If
        If Confirm("Remove all " & n & " note(s) from this block? " & _
                   "What you typed in them is lost.") Then
            modBlock.UngroupParts shp
            modNote.ClearNotes shp
            StyleBlock shp
            Reselect shp
        End If
        Exit Sub
    End If

    modBlock.UngroupParts shp
    If modNote.FindNote(shp, ln) Is Nothing Then modNote.AddNote shp, ln
    StyleBlock shp

    ' Straight into the note's text, the way New block does: the placeholder is
    ' selected, so the first keystroke replaces it.
    Set note = modNote.FindNote(shp, ln)
    If Not note Is Nothing Then
        On Error Resume Next
        note.Select
        note.TextFrame.TextRange.Select
        On Error GoTo 0
        ' A block too wide for a margin puts its notes below it, and the slide
        ' runs out after two or three. Say so rather than leave one hanging off
        ' the bottom edge where it will not be noticed until the lecture.
        If modNote.NoteOffSlide(note) Then
            Warn "There is no room left on this slide for another note. " & _
                 "Make the block smaller, or drag this note where you want it."
        End If
    End If
    Exit Sub
Failed:
    Warn "DoNote failed: " & Err.Description
End Sub

' The three note-style commands share one rule.
'
' Each sets the choice for notes made AFTERWARDS, and applies it to the notes
' you have SINGLED OUT - or, if you have not singled any out, to every note on
' the selected block. Click into a note and pick a colour and only that note
' changes; select the block and pick one and they all do.
'
' They also change only what they are named after. Restyling all three
' properties together, which is what the first version did, is what made every
' note on a slide the same: setting a colour also reset the size and the font,
' so no two notes could differ in anything.

Public Sub DoNoteSize(ByVal pts As Long)
    On Error GoTo Failed
    Dim notes As Collection, blk As Shape, size As Single

    modOptions.SetNoteSize pts
    ' Before the early exit below. The control shows the current choice, so it
    ' has to be re-read whether or not there was anything to restyle.
    RefreshRibbon

    Set notes = NotesToStyle(blk)
    If notes.count = 0 Then Exit Sub

    ' Auto means "from the block", which is only knowable once the block is.
    size = CSng(pts)
    If pts < 1 Then size = modNote.EffectiveNoteSize(blk)

    modNote.CaptureDrags blk
    modBlock.UngroupParts blk
    modNote.ApplyFontSize notes, size
    StyleBlock blk
    Reselect blk
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoNoteSize failed: " & Err.Description
End Sub

Public Sub DoNoteColor(ByVal presetIndex As Long)
    modOptions.SetNoteColor ThemeNotePreset(presetIndex)
    DoNoteColorApply
End Sub

' Paints the current colour onto whatever is singled out. Reached both by
' choosing from the dropdown and by pressing the swatch beside it - the second
' is how you give a second note the colour you gave the first.
Public Sub DoNoteColorApply()
    On Error GoTo Failed
    Dim notes As Collection, blk As Shape

    ' Before the early exit. The swatch shows the colour in use, so it has to be
    ' re-read whether or not there was anything to paint.
    RefreshRibbon

    Set notes = NotesToStyle(blk)
    If notes.count = 0 Then Exit Sub

    ' Colour changes no geometry, so there is nothing to re-place and the
    ' selection can be left exactly where it was - which matters, because
    ' recolouring one note of several is a thing you do repeatedly.
    modNote.ApplyFill notes, modOptions.NoteColor()
    Exit Sub
Failed:
    Warn "DoNoteColor failed: " & Err.Description
End Sub

' Marks the lines the interpreter PRINTED. Everything else in a transcript is
' something you typed, so it gets the prompt - there is no second marking to
' keep in step with this one, and no order of operations that destroys it.
'
' Pressing it on a block that is not yet a transcript makes it one, because
' that is plainly what you meant. With the SHAPE selected rather than text it
' clears the output lines, the way Emphasize and Hide lines do.
Public Sub DoOutputLines()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sel As Selection
    Dim txt As String, a As Long, b As Long, spec As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    spec = modOutput.GetOutputLines(shp)
    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText Then
        txt = shp.TextFrame.TextRange.text
        a = modBlock.LineOfChar(txt, sel.TextRange.Start)
        If sel.TextRange.length > 0 Then
            b = modBlock.LineOfChar(txt, sel.TextRange.Start + sel.TextRange.length - 1)
        Else
            ' A bare cursor marks the one line it sits on.
            b = a
        End If
        ' ADDED to what is already marked: a selection covers one run, and a
        ' transcript has several.
        spec = modOutput.ToggleLines(spec, a, b)
        modOutput.SetTranscript shp, True
    Else
        spec = ""
    End If

    modBlock.UngroupParts shp
    modOutput.SetOutputLines shp, spec
    StyleBlock shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoOutputLines failed: " & Err.Description
End Sub

' Turns the whole block into a session at an interpreter, or back into a plain
' listing. A block-level property rather than a per-line one, which is what
' stops it from fighting the output marking.
Public Sub DoTranscript()
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    modBlock.UngroupParts shp
    modOutput.SetTranscript shp, Not modOutput.IsTranscript(shp)
    StyleBlock shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoTranscript failed: " & Err.Description
End Sub

' Attaches an OUTPUT note to a line: what the program printed, dressed as a
' terminal rather than as an aside.
'
' Same targeting as Note, so there is one answer to "which line do you mean".
' Pressing it on a line that already has a plain note converts that note rather
' than adding a second - two notes cannot share a line, and converting is
' almost certainly what was meant.
Public Sub DoOutputNote()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sel As Selection
    Dim ln As Long, note As Shape, existed As Boolean

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText Then
        ln = modBlock.LineOfChar(shp.TextFrame.TextRange.text, sel.TextRange.Start)
    End If
    If ln < 1 Then ln = modBlock.LastEmphasisedLine(modBlock.GetEmphasis(shp))

    If ln < 1 Then
        Warn "Put the cursor on the line that produced the output, and press " & _
             "Output. On a walkthrough slide the emphasised line is used, so " & _
             "the block itself is enough."
        Exit Sub
    End If

    existed = Not (modNote.FindNote(shp, ln) Is Nothing)
    modBlock.UngroupParts shp
    modNote.AddOutputNote shp, ln
    StyleBlock shp

    Set note = modNote.FindNote(shp, ln)
    If note Is Nothing Then Exit Sub

    On Error Resume Next
    note.Select
    ' Select the text AFTER the mark, so the first keystroke replaces the
    ' placeholder and leaves the arrow where it is.
    If Not existed Then
        note.TextFrame.TextRange.Characters(Len(modNote.OutputMark()) + 1, _
            Len(note.TextFrame.TextRange.text)).Select
    Else
        note.TextFrame.TextRange.Select
    End If
    On Error GoTo 0
    Exit Sub
Failed:
    Warn "DoOutputNote failed: " & Err.Description
End Sub

' Deletes the note you have singled out - the one you have clicked into, or the
' one on the line your cursor is on. Same targeting as the style controls, so
' there is one rule for "which note do you mean" rather than two.
'
' Silent, unlike clearing them all. One note is a small enough loss, and Undo
' takes it back; the all-notes clear on the Note button still asks.
Public Sub DoDeleteNote()
    On Error GoTo Failed
    Dim notes As Collection, blk As Shape, i As Long

    Set notes = NotesToStyle(blk)
    If notes.count = 0 Then
        Warn "Put the cursor on the line whose note you want to delete, or " & _
             "click into the note itself, and press Delete note."
        Exit Sub
    End If

    modNote.CaptureDrags blk
    modBlock.UngroupParts blk
    For i = notes.count To 1 Step -1
        notes(i).Delete
    Next i
    ' StyleBlock clears the leaders and their anchors and redraws what is left,
    ' so the deleted note's connector goes with it.
    StyleBlock blk
    Reselect blk
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoDeleteNote failed: " & Err.Description
End Sub

Public Sub DoNoteFont(ByVal presetIndex As Long)
    On Error GoTo Failed
    Dim notes As Collection, blk As Shape

    modOptions.SetNoteFont ThemeNoteFontValue(presetIndex)
    RefreshRibbon

    ' Deck default cannot be applied to an existing note - Font.Name has no
    ' value meaning "inherit" - so it only affects notes made afterwards.
    If presetIndex = 0 Then Exit Sub

    Set notes = NotesToStyle(blk)
    If notes.count = 0 Then Exit Sub

    modNote.CaptureDrags blk
    modBlock.UngroupParts blk
    modNote.ApplyFontName notes, ThemeNoteFontValue(presetIndex)
    StyleBlock blk
    Reselect blk
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoNoteFont failed: " & Err.Description
End Sub

' The notes a style command should act on, and the block they belong to.
'
' Two ways of singling a note out:
'
'   the note itself is selected      - clicking into the group and picking it
'   the cursor is on its line        - the same gesture that made the note
'
' Anything else - a selected block, a selected group, nothing at all - yields
' NOTHING, and the command only records the choice for the next note.
'
' It used to fall back to "all of the block's notes", which read as a
' convenience and behaved as a trap: the selection is left on the block after
' every other command, so opening the colour list to pick a colour for the NEXT
' note silently repainted the notes already there. Nothing should repaint a note
' unless you have pointed at it.
'
' The cursor route is the one that carries this. Every note ends up inside a
' group with its block, and reaching a shape inside a group takes two deliberate
' clicks, so without it singling a note out would be fiddly enough that nobody
' would do it.
Private Function NotesToStyle(ByRef blk As Shape) As Collection
    Dim c As Collection, shp As Shape, problem As String
    Dim ln As Long, note As Shape

    Set c = SelectedNotes()
    If c.count > 0 Then
        Set blk = modNote.BlockOfNote(c(1))
        If blk Is Nothing Then Set c = New Collection
        Set NotesToStyle = c
        Exit Function
    End If

    Set NotesToStyle = New Collection

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then Exit Function

    ln = CursorLine(shp)
    If ln < 1 Then Exit Function
    Set note = modNote.FindNote(shp, ln)
    If note Is Nothing Then Exit Function

    Set blk = shp
    c.Add note
    Set NotesToStyle = c
End Function

' The line the text cursor is on, or 0 when the selection is not text.
Private Function CursorLine(ByVal shp As Shape) As Long
    Dim sel As Selection
    On Error GoTo Done
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> ppSelectionText Then Exit Function
    CursorLine = modBlock.LineOfChar(shp.TextFrame.TextRange.text, sel.TextRange.Start)
Done:
End Function

' The notes in the current selection.
'
' CHILD shape range first. Clicking into a group and picking one member gives a
' child shape range, while ShapeRange still reports the whole group - so reading
' ShapeRange alone can never see a note, because Stylize groups every note with
' its block. That is not a corner case, it is the normal state of a note.
'
' Late-bound through Object so that HasChildShapeRange, which arrived in
' PowerPoint 2010, cannot turn into a compile error on an older host - the
' add-in claims 2010 and up, and a compile error takes the whole module down
' rather than just this feature.
Private Function SelectedNotes() As Collection
    Dim sel As Object, sr As Object, c As Collection
    Dim i As Long, hasChild As Boolean

    Set c = New Collection
    On Error GoTo Done
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> ppSelectionShapes And sel.Type <> ppSelectionText Then GoTo Done

    On Error Resume Next
    hasChild = sel.HasChildShapeRange
    On Error GoTo Done

    If hasChild Then
        Set sr = sel.ChildShapeRange
    Else
        Set sr = sel.ShapeRange
    End If

    For i = 1 To sr.count
        If modNote.IsNote(sr(i)) Then c.Add sr(i)
    Next i
Done:
    Set SelectedNotes = c
End Function

' Restyles the selected block if there is one, so the change is visible at once.
' Bold is a render-time property, so restyling is all it takes.
Public Sub DoEmphasisBold(ByVal on_ As Boolean)
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    modOptions.SetEmphasisBold on_

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then Exit Sub
    StyleBlock shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoEmphasisBold failed: " & Err.Description
End Sub

' Sets the colour every arrow gets, and repaints the ones already in the deck.
'
' Deck-wide, unlike the note controls. Arrows are uniform by design - two
' colours of arrow on one slide say something the arrows do not mean - so
' "which arrow do you mean" has no answer worth asking, and a walkthrough built
' before you settled on a colour would otherwise need recolouring slide by slide.
Public Sub DoArrowColor(ByVal presetIndex As Long)
    On Error GoTo Failed
    Dim pres As Presentation

    modOptions.SetArrowColor ThemeArrowPreset(presetIndex)
    RefreshRibbon

    On Error Resume Next
    Set pres = Application.ActivePresentation
    On Error GoTo Failed
    If pres Is Nothing Then Exit Sub

    modArrow.RecolorAll pres, ThemeArrowPreset(presetIndex)
    Exit Sub
Failed:
    Warn "DoArrowColor failed: " & Err.Description
End Sub

' Puts a block arrow in the left margin beside one line, or takes it away again.
'
' The other half of Emphasize. Emphasize fades everything else; this leaves the
' code whole and points at a line from outside it, which is what you want when
' the whole snippet has to stay readable.
'
' Same targeting as Note, so there is one answer to "which line do you mean":
' the cursor's line, or - with the block itself selected - the emphasised one,
' which is what makes it one click per slide on a generated walkthrough. With
' the block selected and nothing emphasised, it clears the arrows. No
' confirmation for that, unlike notes: an arrow holds nothing you typed.
Public Sub DoArrow()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sel As Selection, ln As Long

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText Then
        ln = modBlock.LineOfChar(shp.TextFrame.TextRange.text, sel.TextRange.Start)
    End If
    If ln < 1 Then ln = modBlock.LastEmphasisedLine(modBlock.GetEmphasis(shp))

    modBlock.UngroupParts shp
    If ln < 1 Then
        If modArrow.ArrowCount(shp) = 0 Then
            Warn "Put the cursor on the line you want to point at, and press " & _
                 "Arrow. On a walkthrough slide the emphasised line is used, so " & _
                 "the block itself is enough."
            modBlock.GroupParts shp
            Exit Sub
        End If
        modArrow.ClearArrows shp
    ' Output marking goes too: it is a rendering choice about lines, like the
    ' emphasis and the covers, and Stylize brings it all back from the tag.
    modOutput.SetOutputLines shp, ""
    modOutput.SetTranscript shp, False
    ElseIf Not modArrow.RemoveArrow(shp, ln) Then
        modArrow.AddArrow shp, ln
    End If

    StyleBlock shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoArrow failed: " & Err.Description
End Sub

' Duplicates the slide with nothing hidden, so the answer follows the question.
Public Sub DoReveal()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sld As Slide, dup As Slide, target As Shape

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    If Len(modBlock.GetHidden(shp)) = 0 Then
        Warn "Nothing is hidden on this block. Select some lines and press Hide first."
        Exit Sub
    End If

    Set sld = modGutter.OwningSlide(shp)
    Set dup = sld.Duplicate(1)
    dup.MoveTo sld.SlideIndex + 1

    Set target = BlockOnSlide(dup, shp.Tags(modBlock.TAG_ID))
    If Not target Is Nothing Then
        modBlock.SetHidden target, ""
        StyleBlock target
    End If

    sld.Select
    Exit Sub
Failed:
    Warn "DoReveal failed: " & Err.Description
End Sub

Public Sub DoStepThrough(ByVal cumulative As Boolean)
    On Error GoTo Failed
    Dim shp As Shape, problem As String, sld As Slide, dup As Slide, target As Shape
    Dim txt As String, lines() As String, sel As Selection
    Dim steps() As Long, n As Long, i As Long, k As Long
    Dim firstLine As Long, lastLine As Long, base As Long, list As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    Set sld = modGutter.OwningSlide(shp)
    txt = shp.TextFrame.TextRange.text
    lines = modBlock.SplitLines(txt)

    ' A text selection limits the walk to those lines. Without it a forty-line
    ' block would produce forty slides when you wanted to walk one function.
    firstLine = 1
    lastLine = UBound(lines) - LBound(lines) + 1
    Set sel = Application.ActiveWindow.Selection
    If sel.Type = ppSelectionText Then
        If sel.TextRange.length > 0 Then
            firstLine = modBlock.LineOfChar(txt, sel.TextRange.Start)
            lastLine = modBlock.LineOfChar(txt, sel.TextRange.Start + sel.TextRange.length - 1)
        End If
    End If

    ' Blank lines get no slide of their own - a dead beat in a walkthrough.
    ReDim steps(1 To lastLine - firstLine + 1)
    For i = firstLine To lastLine
        If Len(Trim$(lines(LBound(lines) + i - 1))) > 0 Then
            n = n + 1
            steps(n) = i
        End If
    Next i

    If n = 0 Then
        Warn "No code lines to step through."
        Exit Sub
    End If

    If n > 12 Then
        If Not Confirm("This will add " & n & " slides after this one." & vbCrLf & _
                       "Carry on?") Then Exit Sub
    End If

    ' The starting slide shows the code with nothing picked out.
    modBlock.SetEmphasis shp, ""
    StyleBlock shp

    base = sld.SlideIndex
    For k = 1 To n
        Set dup = sld.Duplicate(1)
        ' Duplicates land immediately after the original, so without this they
        ' would come out in reverse order.
        dup.MoveTo base + k

        Set target = BlockOnSlide(dup, shp.Tags(modBlock.TAG_ID))
        If Not target Is Nothing Then
            If cumulative Then
                list = ""
                For i = 1 To k
                    If Len(list) > 0 Then list = list & ","
                    list = list & CStr(steps(i))
                Next i
            Else
                list = CStr(steps(k))
            End If
            ' Arrows INSTEAD of emphasis, when the deck asks for it. Step
            ' through moves one arrow down the margin; Build up leaves them
            ' behind, so the marks accumulate the way the code does.
            If modOptions.StepArrow() Then
                modBlock.SetEmphasis target, ""
                If cumulative Then
                    For i = 1 To k
                        modArrow.AddArrow target, steps(i)
                    Next i
                Else
                    modArrow.ClearArrows target
                    modArrow.AddArrow target, steps(k)
                End If
            Else
                modBlock.SetEmphasis target, list
            End If
            ' A note per step, already attached to the line this slide is
            ' about, so the walkthrough arrives ready to be written into.
            ' steps(k) is the newly emphasised line either way - for Build up
            ' it is the last of the range, which is where the slide has reached.
            If modOptions.StepNote() Then
                If modNote.FindNote(target, steps(k)) Is Nothing Then
                    modNote.AddNote target, steps(k)
                End If
            End If
            StyleBlock target
        End If
    Next k

    ' Spotlight ends on a clean slide, so the walk finishes by handing the whole
    ' code back rather than leaving the last line singled out. Build up already
    ' ends with everything lit, so it needs no such slide.
    If Not cumulative Then
        Set dup = sld.Duplicate(1)
        dup.MoveTo base + n + 1
        Set target = BlockOnSlide(dup, shp.Tags(modBlock.TAG_ID))
        If Not target Is Nothing Then
            modBlock.SetEmphasis target, ""
            StyleBlock target
        End If
    End If

    sld.Select
    Exit Sub
Failed:
    Warn "DoStepThrough failed: " & Err.Description
End Sub

' The block carrying this id on a given slide. Duplicated slides carry the same
' id, which is fine - the search never crosses a slide.
Private Function BlockOnSlide(ByVal sld As Slide, ByVal blockId As String) As Shape
    Dim shp As Shape
    For Each shp In modBlock.AllShapes(sld)
        If shp.Tags(modBlock.TAG_BLOCK) = "1" Then
            If shp.Tags(modBlock.TAG_ID) = blockId Then
                Set BlockOnSlide = shp
                Exit Function
            End If
        End If
    Next shp
End Function

Public Sub DoSizeUp()
    StepSize 1
End Sub

Public Sub DoSizeDown()
    StepSize -1
End Sub

' Picks the largest rung the code fits at, and says so when that is too small
' to read from the back of a room rather than silently shrinking it.
Public Sub DoFit()
    On Error GoTo Failed
    Dim shp As Shape, problem As String, best As Single, slides As Long

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    best = FitSizeFor(shp)
    Resize shp, best

    If best < modSpec.MIN_TEACHING_SIZE Then
        slides = SlidesNeededFor(shp)
        Warn "This only fits at " & Format$(best, "0") & "pt, which will not read " & _
             "from the back of a lecture hall." & vbCrLf & vbCrLf & _
             "At " & Format$(modSpec.MIN_TEACHING_SIZE, "0") & "pt it needs about " & _
             slides & " slide" & IIf(slides = 1, "", "s") & ". Consider splitting it."
    End If
    Exit Sub
Failed:
    Warn "DoFit failed: " & Err.Description
End Sub

' Larger is CAPPED at the fitting size. The lab showed what happens without
' that guard: at 32pt a long line clipped off the right edge and the block
' overran the top. Growth stops where the content stops fitting, not where the
' ladder ends.
Private Sub StepSize(ByVal direction As Long)
    Dim shp As Shape, problem As String
    Dim idx As Long, newSize As Single, capSize As Single

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    idx = modSpec.LadderIndexOf(modBlock.BlockFontSize(shp)) + direction
    If idx < 0 Then idx = 0
    If idx > modSpec.LadderCount() - 1 Then idx = modSpec.LadderCount() - 1
    newSize = modSpec.LadderAt(idx)

    If direction > 0 Then
        capSize = FitSizeFor(shp)
        If newSize > capSize Then
            Warn "Already as large as this block can go and still fit the slide."
            Exit Sub
        End If
    End If

    Resize shp, newSize
End Sub

' Changing size RESETS every run to the default colour, because the font has to
' be reapplied across the whole range. So re-highlighting is not a nicety here,
' it is the other half of the operation.
Private Sub Resize(ByVal shp As Shape, ByVal newSize As Single)
    ' Before ApplySize, which moves the block. StyleBlock's own capture then
    ' becomes a no-op - modNote's header says why it has to.
    modNote.CaptureDrags shp
    modBlock.UngroupParts shp
    modBlock.ApplySize shp, newSize
    ' The rest is the ordinary pipeline. It used to be copied out here, and the
    ' copy fell behind: a resized block lost its covers and its notes.
    StyleBlock shp
    Reselect shp
    RefreshRibbon
End Sub

Private Function FitSizeFor(ByVal shp As Shape) As Single
    Dim txt As String
    txt = shp.TextFrame.TextRange.text
    FitSizeFor = modSpec.SpecFitSize(modBlock.CountLines(txt), _
                                     modBlock.LongestLine(txt), HasGutter(shp))
End Function

Private Function SlidesNeededFor(ByVal shp As Shape) As Long
    Dim txt As String
    txt = shp.TextFrame.TextRange.text
    SlidesNeededFor = modSpec.SpecSlidesNeeded(modBlock.CountLines(txt), _
                                               modBlock.LongestLine(txt), HasGutter(shp))
End Function

Private Function HasGutter(ByVal shp As Shape) As Boolean
    HasGutter = modGutter.HasGutter(shp)
End Function

Public Sub DoSetLanguage(ByVal index As Long)
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    mLangIndex = index

    ' If a block is selected, the choice applies to it and takes effect at once.
    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then Exit Sub

    modBlock.SetBlockLang shp, CurrentLangId()
    modRender.ApplyHighlight shp, CurrentLangId()
    Exit Sub
Failed:
    Warn "DoSetLanguage failed: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Ribbon callbacks
'------------------------------------------------------------------------------

Public Sub RibbonNewBlock(control As IRibbonControl)
    DoNewBlock
End Sub

Public Sub RibbonStylize(control As IRibbonControl)
    DoStylize
End Sub

Public Sub RibbonLangCount(control As IRibbonControl, ByRef count)
    count = modLangRegistry.LangCount()
End Sub

Public Sub RibbonLangLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = modLangRegistry.LangAt(CLng(index)).DisplayName
End Sub

Public Sub RibbonLangSelected(control As IRibbonControl, ByRef index)
    index = mLangIndex
End Sub

Public Sub RibbonLangChanged(control As IRibbonControl, id As String, index As Integer)
    DoSetLanguage CLng(index)
End Sub

Public Sub RibbonGutterPressed(control As IRibbonControl, ByRef returnedVal)
    Dim shp As Shape, problem As String
    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        returnedVal = False
    Else
        returnedVal = modGutter.HasGutter(shp)
    End If
End Sub

Public Sub RibbonToggleGutter(control As IRibbonControl, pressed As Boolean)
    DoToggleGutter
End Sub

Public Sub RibbonFirstLineText(control As IRibbonControl, ByRef text)
    Dim shp As Shape, problem As String
    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        text = "1"
    Else
        text = CStr(modGutter.FirstLine(shp))
    End If
End Sub

Public Sub RibbonFirstLineChanged(control As IRibbonControl, text As String)
    Dim n As Long
    n = CLng(Val(text))
    If n < 1 Or n > 99999 Then
        Warn "Enter the number the first line should get, between 1 and 99999."
        RefreshRibbon
        Exit Sub
    End If
    DoFirstLine n
End Sub

Public Sub RibbonGuidesPressed(control As IRibbonControl, ByRef returnedVal)
    Dim shp As Shape, problem As String
    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        returnedVal = False         ' the default for a new block
    Else
        returnedVal = modGuides.GuidesEnabled(shp)
    End If
End Sub

Public Sub RibbonToggleGuides(control As IRibbonControl, pressed As Boolean)
    DoToggleGuides
End Sub

' The size box: shows the selected block's size, and accepts a typed one.
Public Sub RibbonHide(control As IRibbonControl)
    DoHide
End Sub

Public Sub RibbonReveal(control As IRibbonControl)
    DoReveal
End Sub

Public Sub RibbonNote(control As IRibbonControl)
    DoNote
End Sub

' Item 0 is Auto - a note sized from the block it belongs to, which is the
' default and what almost every deck wants. The rest is the ordinary ladder.
Public Sub RibbonNoteSizeCount(control As IRibbonControl, ByRef count)
    count = modSpec.LadderCount() + 1
End Sub

Public Sub RibbonNoteSizeLabel(control As IRibbonControl, index As Integer, ByRef label)
    If index = 0 Then
        label = "Auto"
    Else
        label = Format$(modSpec.LadderAt(CLng(index) - 1), "0")
    End If
End Sub

' Shows the SELECTED note's size when there is one, and the deck's choice
' otherwise. Same for the colour and the font below.
'
' PowerPoint gives an add-in no selection-changed event, so these are re-read
' after a command rather than as you click about - the language dropdown has
' always had the same limitation. They are correct whenever the add-in has just
' done something, which is when they are looked at.
Public Sub RibbonNoteSizeText(control As IRibbonControl, ByRef text)
    Dim note As Shape, pts As Long

    Set note = FirstSelectedNote()
    If Not note Is Nothing Then
        text = Format$(note.TextFrame.TextRange.Font.size, "0")
        Exit Sub
    End If

    pts = modOptions.NoteSize()
    If pts < 1 Then
        text = "Auto"
    Else
        text = CStr(pts)
    End If
End Sub

Public Sub RibbonNoteSizeChanged(control As IRibbonControl, text As String)
    Dim want As Long

    If LCase$(Trim$(text)) = "auto" Then
        DoNoteSize modOptions.NOTE_SIZE_AUTO
        RefreshRibbon
        Exit Sub
    End If

    want = CLng(Val(text))
    If want < 6 Or want > 96 Then
        Warn "Enter a size between 6 and 96 points, or Auto to size notes " & _
             "from the block."
        RefreshRibbon
        Exit Sub
    End If
    DoNoteSize want
    RefreshRibbon
End Sub

Public Sub RibbonNoteColorCount(control As IRibbonControl, ByRef count)
    count = ThemeNotePresetCount()
End Sub

Public Sub RibbonNoteColorLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = ThemeNotePresetName(CLng(index))
End Sub

Public Sub RibbonNoteColorSelected(control As IRibbonControl, ByRef index)
    index = ThemeNotePresetIndexOf(CurrentNoteColor())
End Sub

Public Sub RibbonNoteColorChanged(control As IRibbonControl, id As String, index As Integer)
    DoNoteColor CLng(index)
End Sub

' A swatch in the actual colour, so the list shows what you are choosing rather
' than asking you to remember what "Plum" looked like.
Public Sub RibbonNoteColorImage(control As IRibbonControl, index As Integer, ByRef image)
    Dim p As Object
    Set p = modSwatch.Swatch(ThemeNotePreset(CLng(index)))
    ' No image is a working gallery with labels only. A failed swatch must not
    ' take the control down with it.
    If Not p Is Nothing Then Set image = p
End Sub

' The swatch shown beside the closed dropdown.
'
' A dropDown displays its selected item's LABEL and nothing else - the item
' images only appear in the open list - and unlike a gallery it has no image of
' its own. So the colour in use is shown by a button sitting next to it, which
' also gives it something to do: pressing it applies that colour again, which is
' how you give a second note the same colour as the first.
Public Sub RibbonNoteColorFace(control As IRibbonControl, ByRef image)
    Dim p As Object
    Set p = modSwatch.Swatch(CurrentNoteColor())
    If Not p Is Nothing Then Set image = p
End Sub

Public Sub RibbonNoteColorApply(control As IRibbonControl)
    DoNoteColorApply
End Sub

Public Sub RibbonOutputLines(control As IRibbonControl)
    DoOutputLines
End Sub

Public Sub RibbonTranscriptPressed(control As IRibbonControl, ByRef returnedVal)
    Dim shp As Shape, problem As String
    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        returnedVal = False
    Else
        returnedVal = modOutput.IsTranscript(shp)
    End If
End Sub

Public Sub RibbonToggleTranscript(control As IRibbonControl, pressed As Boolean)
    DoTranscript
End Sub

Public Sub RibbonOutputNote(control As IRibbonControl)
    DoOutputNote
End Sub

Public Sub RibbonDeleteNote(control As IRibbonControl)
    DoDeleteNote
End Sub

Public Sub RibbonArrow(control As IRibbonControl)
    DoArrow
End Sub

Public Sub RibbonArrowColorCount(control As IRibbonControl, ByRef count)
    count = ThemeArrowPresetCount()
End Sub

Public Sub RibbonArrowColorLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = ThemeArrowPresetName(CLng(index))
End Sub

Public Sub RibbonArrowColorSelected(control As IRibbonControl, ByRef index)
    index = ThemeArrowPresetIndexOf(modOptions.ArrowColor())
End Sub

Public Sub RibbonArrowColorChanged(control As IRibbonControl, id As String, index As Integer)
    DoArrowColor CLng(index)
End Sub

Public Sub RibbonArrowColorImage(control As IRibbonControl, index As Integer, ByRef image)
    Dim p As Object
    Set p = modSwatch.Swatch(ThemeArrowPreset(CLng(index)))
    If Not p Is Nothing Then Set image = p
End Sub

' The swatch beside the closed list, for the same reason the note one exists:
' a dropDown shows its selected item's label and has no image of its own.
Public Sub RibbonArrowColorFace(control As IRibbonControl, ByRef image)
    Dim p As Object
    Set p = modSwatch.Swatch(modOptions.ArrowColor())
    If Not p Is Nothing Then Set image = p
End Sub

Public Sub RibbonArrowColorApply(control As IRibbonControl)
    DoArrowColor ThemeArrowPresetIndexOf(modOptions.ArrowColor())
End Sub

Public Sub RibbonStepArrowPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = modOptions.StepArrow()
End Sub

Public Sub RibbonToggleStepArrow(control As IRibbonControl, pressed As Boolean)
    modOptions.SetStepArrow pressed
    RefreshRibbon
End Sub

Private Function CurrentNoteColor() As Long
    Dim note As Shape
    Set note = FirstSelectedNote()
    If note Is Nothing Then
        CurrentNoteColor = modOptions.NoteColor()
    Else
        On Error Resume Next
        CurrentNoteColor = modOptions.NoteColor()
        CurrentNoteColor = note.fill.ForeColor.RGB
    End If
End Function

' The one note the ribbon reports on. Nothing when the selection is not a note.
Private Function FirstSelectedNote() As Shape
    Dim c As Collection
    Set c = SelectedNotes()
    If c.count > 0 Then Set FirstSelectedNote = c(1)
End Function

Public Sub RibbonNoteFontCount(control As IRibbonControl, ByRef count)
    count = ThemeNoteFontCount()
End Sub

Public Sub RibbonNoteFontLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = ThemeNoteFontName(CLng(index))
End Sub

Public Sub RibbonNoteFontSelected(control As IRibbonControl, ByRef index)
    Dim note As Shape
    Set note = FirstSelectedNote()
    If note Is Nothing Then
        index = ThemeNoteFontIndexOf(modOptions.NoteFont())
    Else
        On Error Resume Next
        index = ThemeNoteFontIndexOf(note.TextFrame.TextRange.Font.Name)
    End If
End Sub

Public Sub RibbonNoteFontChanged(control As IRibbonControl, id As String, index As Integer)
    DoNoteFont CLng(index)
End Sub

Public Sub RibbonStepNotePressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = modOptions.StepNote()
End Sub

Public Sub RibbonToggleStepNote(control As IRibbonControl, pressed As Boolean)
    modOptions.SetStepNote pressed
    RefreshRibbon
End Sub

Public Sub RibbonEmphasisBoldPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = modOptions.EmphasisBold()
End Sub

Public Sub RibbonToggleEmphasisBold(control As IRibbonControl, pressed As Boolean)
    DoEmphasisBold pressed
End Sub

Public Sub RibbonStepThrough(control As IRibbonControl)
    DoStepThrough False
End Sub

Public Sub RibbonBuildUp(control As IRibbonControl)
    DoStepThrough True
End Sub

Public Sub RibbonEmphasize(control As IRibbonControl)
    DoEmphasize
End Sub

Public Sub RibbonCopyCode(control As IRibbonControl)
    DoCopyCode
End Sub

Public Sub RibbonStrip(control As IRibbonControl)
    DoStrip
End Sub

Public Sub RibbonSizeCount(control As IRibbonControl, ByRef count)
    count = modSpec.LadderCount()
End Sub

Public Sub RibbonSizeLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = Format$(modSpec.LadderAt(CLng(index)), "0")
End Sub

Public Sub RibbonSizeText(control As IRibbonControl, ByRef text)
    Dim shp As Shape, problem As String
    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        text = Format$(modSpec.BASE_SIZE, "0")
    Else
        text = Format$(modBlock.BlockFontSize(shp), "0")
    End If
End Sub

' Accepts any size, not only a rung - typing 23 is a reasonable thing to do.
' Still capped at what fits, for the same reason Larger is.
Public Sub RibbonSizeChanged(control As IRibbonControl, text As String)
    Dim shp As Shape, problem As String, want As Single, capSize As Single

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    want = Val(text)
    If want < 6 Or want > 96 Then
        Warn "Enter a size between 6 and 96 points."
        RefreshRibbon
        Exit Sub
    End If

    capSize = FitSizeFor(shp)
    If want > capSize Then
        Warn "At " & Format$(want, "0") & "pt this block would not fit the slide." & _
             vbCrLf & "The largest that fits is " & Format$(capSize, "0") & "pt."
        RefreshRibbon
        Exit Sub
    End If

    Resize shp, want
End Sub

Public Sub RibbonSizeUp(control As IRibbonControl)
    DoSizeUp
End Sub

Public Sub RibbonSizeDown(control As IRibbonControl)
    DoSizeDown
End Sub

Public Sub RibbonFit(control As IRibbonControl)
    DoFit
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------

Public Function CurrentLangId() As String
    CurrentLangId = modLangRegistry.LangAt(mLangIndex).id
End Function

' Quiet mode exists for the test harness. A modal warning in an automated run
' blocks until someone clicks OK, which is indistinguishable from a hang.
Public Sub SetQuiet(ByVal quiet As Boolean)
    mQuiet = quiet
End Sub

Public Function LastWarning() As String
    LastWarning = mLastWarning
End Function

' Yes/no, and always yes in a test run - a modal question would hang it.
Private Function Confirm(ByVal msg As String) As Boolean
    If mQuiet Then
        Confirm = True
        Exit Function
    End If
    Confirm = (MsgBox(msg, vbQuestion + vbYesNo, ADDIN_NAME) = vbYes)
End Function

Private Sub Warn(ByVal msg As String)
    mLastWarning = msg
    If Not mQuiet Then MsgBox msg, vbExclamation, ADDIN_NAME
End Sub

Private Sub NotYet(ByVal what As String)
    Warn what & " is not built yet."
End Sub

' Puts the selection back on the block after a command.
'
' Without this every action costs a re-selection: press Larger, lose the
' selection, click the block, press Larger again. Worse, it makes the add-in
' look broken - you type a line, press Stylize, and nothing happens, because
' nothing was selected any more.
'
' Selects the GROUP when there is one, so the next drag moves the whole thing.
Private Sub Reselect(ByVal shp As Shape)
    Dim g As Shape
    On Error Resume Next
    Set g = modBlock.ParentGroup(shp)
    If g Is Nothing Then
        shp.Select
    Else
        g.Select
    End If
    On Error GoTo 0
End Sub

Private Function ActiveSlide() As Slide
    ' Fails in slide sorter and in any view with no current slide, so the
    ' caller checks for Nothing rather than trusting it.
    On Error Resume Next
    Set ActiveSlide = Application.ActiveWindow.View.Slide
    On Error GoTo 0
End Function

' The comment marker comes from the language table, so this line is not Python
' specific even though it looks it.
Private Function PlaceholderFor(ByRef lang As LangDef) As String
    Dim parts() As String, marker As String
    marker = ""
    If Len(lang.LineComments) > 0 Then
        parts = Split(Trim$(lang.LineComments), " ")
        marker = parts(0) & " "
    End If
    PlaceholderFor = marker & "type your code here"
End Function
