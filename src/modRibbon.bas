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
    modBlock.UngroupParts shp
    modRender.ApplyHighlight shp, langId
    modBlock.ResizeToContent shp
    ' Gutter before guides: it changes the left margin, and the guides are
    ' placed from that margin.
    modGutter.SyncGutter shp
    modGuides.DrawGuides shp
    ' Back into a group, so the block and its numbers and guides drag as one.
    modBlock.GroupParts shp
End Sub

Public Sub DoStylizeAll()
    On Error GoTo Failed
    Dim sld As Slide, shp As Shape, count As Long

    Set sld = ActiveSlide()
    If sld Is Nothing Then
        Warn "Open a slide in Normal view first."
        Exit Sub
    End If

    For Each shp In modBlock.AllShapes(sld)
        If modBlock.IsCodeBlock(shp) And shp.HasTextFrame Then
            modBlock.UngroupParts shp
            modRender.ApplyHighlight shp, modBlock.BlockLangId(shp, CurrentLangId())
            modBlock.ResizeToContent shp
            modGutter.SyncGutter shp
            modGuides.DrawGuides shp
            modBlock.GroupParts shp
            count = count + 1
        End If
    Next shp

    If count = 0 Then
        Warn "No code blocks on this slide. Stylize one first, which tags it."
    End If
    Exit Sub
Failed:
    Warn "DoStylizeAll failed: " & Err.Description
End Sub

Public Sub DoToggleGutter()
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    modBlock.UngroupParts shp
    modGutter.ToggleGutter shp
    modBlock.ResizeToContent shp
    modGutter.SyncGutter shp
    modGuides.DrawGuides shp
    modBlock.GroupParts shp
    Reselect shp
    RefreshRibbon
    Exit Sub
Failed:
    Warn "DoToggleGutter failed: " & Err.Description
End Sub

Public Sub DoToggleGuides()
    On Error GoTo Failed
    Dim shp As Shape, problem As String

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    modBlock.UngroupParts shp
    modGuides.SetGuidesEnabled shp, Not modGuides.GuidesEnabled(shp)
    modGuides.DrawGuides shp
    modBlock.GroupParts shp
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
    shp.TextFrame.TextRange.Copy
    Exit Sub
Failed:
    Warn "DoCopyCode failed: " & Err.Description
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

    modBlock.ResizeToContent shp
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
            modBlock.SetEmphasis target, list
            StyleBlock target
        End If
    Next k

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
    modBlock.UngroupParts shp
    modBlock.ApplySize shp, newSize
    modRender.ApplyHighlight shp, modBlock.BlockLangId(shp, CurrentLangId())
    modGutter.SyncGutter shp
    modGuides.DrawGuides shp
    modBlock.GroupParts shp
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

Public Sub RibbonStylizeAll(control As IRibbonControl)
    DoStylizeAll
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
