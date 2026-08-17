Attribute VB_Name = "modSelfTest"
'==============================================================================
' modSelfTest - runs the scanner over the sample corpus and dumps what it
'               decided, so it can be diffed against tools/lexref.py.
'
' This module exists to make the lexer testable WITHOUT PowerPoint in the loop
' as a human step. tools/run-lexer-tests.sh drives it over COM: it imports the
' modules into a scratch deck, calls MaskFileToFile once per sample, and diffs
' the result against the reference. A lexer bug shows up as a wrong column in a
' diff rather than as a vague sense that the colours look off.
'==============================================================================
Option Explicit

' The language for this run. Set once with SetLanguage rather than passed on
' every call, because Application.Run proved unreliable with three arguments
' from PowerShell - two always work, three sometimes report "Sub or function
' not defined" even though the procedure is plainly there. Not root-caused;
' the dependency is simply removed.
Private mLangId As String

Public Sub SetLanguage(ByVal langId As String)
    mLangId = langId
End Sub

Private Function RunLangId() As String
    If Len(mLangId) = 0 Then mLangId = modLangRegistry.DefaultLangId()
    RunLangId = mLangId
End Function

Private Function RunLang() As LangDef
    If Len(mLangId) = 0 Then mLangId = modLangRegistry.DefaultLangId()
    RunLang = modLangRegistry.GetLang(mLangId)
End Function

' Smoke test for the harness itself. VBA compiles on demand, so a module that
' fails to compile reports as "Sub or function not defined" when you call into
' it - which reads exactly like a wrong macro name. Calling this first separates
' the two: if Ping works and the rest do not, the module is broken, not the
' plumbing.
Public Function Ping() As String
    Ping = "pong " & modLangRegistry.LangCount() & " lang(s)"
End Function

' Masks EVERY sample in a folder in a single call, and returns a summary.
'
' One macro call for the whole corpus rather than three per sample. Application.Run
' from PowerShell proved unreliable when called repeatedly against a freshly
' imported project - procedures that were demonstrably present reported as
' "Sub or function not defined". Doing the whole run inside VBA sidesteps it, and
' is faster anyway: the COM boundary is crossed once instead of thirty times.
'
' Returns one line per sample: "<name> <maskLen> <spanCount>", or
' "<name> ERROR <description>".
Public Function RunCorpus(ByVal sampleDir As String, ByVal outDir As String) As String
    Dim names() As String, count As Long, i As Long
    Dim fileName As String, src As String, mask As String, report As String
    Dim spans() As Span, n As Long
    Dim t0 As Single, totalChars As Long

    ReDim names(0 To 255)
    count = 0

    ' Collect the names first. Interleaving Dir() with the file reads below
    ' would clobber its iteration state.
    fileName = Dir(sampleDir & "\*.py")
    Do While Len(fileName) > 0
        If count > UBound(names) Then ReDim Preserve names(0 To count + 255)
        names(count) = fileName
        count = count + 1
        fileName = Dir()
    Loop

    t0 = Timer
    For i = 0 To count - 1
        On Error GoTo ItemFailed
        src = ReadTextFile(sampleDir & "\" & names(i))
        mask = modLexer.MaskOf(src, RunLang())
        n = modLexer.Tokenize(src, RunLang(), spans)
        WriteTextFile outDir & "\" & BaseName(names(i)) & ".mask", mask
        report = report & names(i) & " " & Len(mask) & " " & n & vbLf
        totalChars = totalChars + Len(src)
        GoTo NextItem
ItemFailed:
        report = report & names(i) & " ERROR " & Err.Description & vbLf
        Err.Clear
NextItem:
        On Error GoTo 0
    Next i

    ' Timing rides along here rather than in its own macro call. Repeated
    ' Application.Run calls against a freshly imported project proved unreliable,
    ' so the whole run is kept to as few crossings as possible.
    report = report & "TOTAL " & CLng((Timer - t0) * 1000) & " " & totalChars & vbLf

    RunCorpus = report
End Function

Private Function BaseName(ByVal fileName As String) As String
    Dim dot As Long
    dot = InStrRev(fileName, ".")
    If dot > 0 Then BaseName = Left$(fileName, dot - 1) Else BaseName = fileName
End Function

' THE THIN SLICE: insert one block on a fresh slide, highlight it, export a PNG.
'
' Everything in one macro call, both because repeated Application.Run calls
' proved unreliable (section 4 of HANDOFF.md) and because the timings are only
' meaningful measured inside VBA.
'
' Also writes the mask of the SHAPE'S OWN TEXT next to the PNG. That is the
' check that matters here: if PowerPoint altered the text on the way in - line
' endings, tabs, autocorrected quotes - the shape mask and the file mask
' disagree, and every colour is shifted. Comparing renders by eye would not
' catch a one-character shift, but a diff does.
Public Function BuildSlice(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim code As String, shapeText As String, report As String
    Dim t0 As Single, buildMs As Long, colourMs As Long
    Dim applied As Long, dot As Long, guides As Long

    On Error GoTo Failed

    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)

    t0 = Timer
    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, RunLangId())
    buildMs = CLng((Timer - t0) * 1000)

    t0 = Timer
    applied = modRender.ApplyHighlight(shp, RunLangId())
    colourMs = CLng((Timer - t0) * 1000)
    guides = modGuides.DrawGuides(shp)

    ' The mask of what is actually in the shape, not of what was in the file.
    shapeText = shp.TextFrame.TextRange.text
    dot = InStrRev(pngPath, ".")
    If dot > 0 Then
        WriteTextFile Left$(pngPath, dot - 1) & ".mask", _
                      modLexer.MaskOf(shapeText, RunLang())
    End If

    sld.Export pngPath, "PNG", 1920, 1080

    BuildSlice = "lines=" & modBlock.CountLines(shapeText) & _
                 " shapechars=" & Len(shapeText) & _
                 " filechars=" & Len(code) & _
                 " applied=" & applied & _
                 " build_ms=" & buildMs & _
                 " colour_ms=" & colourMs & _
                 " guides=" & guides & _
                 " height=" & Format$(shp.Height, "0.0") & _
                 " top=" & Format$(shp.Top, "0.0")
    Exit Function

Failed:
    BuildSlice = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Drives the REAL ribbon command path - DoNewBlock, DoStylize
' and the wrong-selection cases - not the helpers underneath it. That is the
' point: a command whose logic lives inside an IRibbonControl callback can only
' be reached with a mouse, and would go untested.
'
' Returns "key=value" pairs. Anything ending _ok is an assertion.
Public Function RibbonSliceTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim code As String, r As String
    Dim wantH As Single, gotH As Single, blocks As Long

    On Error GoTo Failed
    modRibbon.SetQuiet True          ' a modal warning would hang the run

    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    ' --- nothing selected -----------------------------------------------------
    Application.ActiveWindow.Selection.Unselect
    modRibbon.DoStylize
    r = r & "warn_empty_selection=" & Quoted(modRibbon.LastWarning()) & vbLf

    ' --- New block ------------------------------------------------------------
    modRibbon.DoNewBlock
    If sld.Shapes.count <> 1 Then
        RibbonSliceTest = r & "newblock_ok=0 shapes=" & sld.Shapes.count
        Exit Function
    End If
    Set shp = sld.Shapes(1)
    r = r & "newblock_ok=" & Abs(CLng(modBlock.IsCodeBlock(shp))) & vbLf
    r = r & "newblock_lang=" & shp.Tags(modBlock.TAG_LANG) & vbLf

    ' --- an empty shape is not a code block -----------------------------------
    ' An oval DOES have a text frame, so this exercises the empty-text guard
    ' rather than the no-text-frame one. The no-text-frame path (a picture) is
    ' still unexercised - it needs an image file to construct.
    sld.Shapes.AddShape(msoShapeOval, 10, 10, 40, 40).Select
    modRibbon.DoStylize
    r = r & "warn_empty_shape=" & Quoted(modRibbon.LastWarning()) & vbLf
    sld.Shapes(sld.Shapes.count).Delete

    ' --- a group is not a code block ------------------------------------------
    sld.Shapes.AddShape(msoShapeOval, 10, 10, 40, 40).Name = "g1"
    sld.Shapes.AddShape(msoShapeOval, 60, 10, 40, 40).Name = "g2"
    sld.Shapes.Range(Array("g1", "g2")).Group.Select
    modRibbon.DoStylize
    r = r & "warn_group=" & Quoted(modRibbon.LastWarning()) & vbLf
    sld.Shapes(sld.Shapes.count).Delete

    ' --- type into the block, then Stylize ----------------------------------
    shp.TextFrame.TextRange.text = modBlock.NormalizeParagraphs(code)
    shp.Select
    modRibbon.DoStylize

    wantH = modSpec.SpecHeight(modBlock.BlockFontSize(shp), _
                               modBlock.CountLines(shp.TextFrame.TextRange.text))
    gotH = shp.Height
    r = r & "lines=" & modBlock.CountLines(shp.TextFrame.TextRange.text) & vbLf
    ' Autofit owns the height now, and computes it a few points tighter than
    ' SpecHeight because it adds no line spacing below the last line. So this
    ' checks the block HUGS - within one line of the predicted height - rather
    ' than matching it exactly.
    r = r & "height_want=" & Format$(wantH, "0.0") & vbLf
    r = r & "height_got=" & Format$(gotH, "0.0") & vbLf
    r = r & "hugs_ok=" & Abs(CLng(Abs(gotH - wantH) <= modSpec.SpecLine(modSpec.BASE_SIZE))) & vbLf
    r = r & "width=" & Format$(shp.Width, "0.0") & _
            " (content " & Format$(modSpec.CONTENT_W, "0.0") & ")" & vbLf
    ' Position matters as much as height. A block can be exactly the right size
    ' and still hang off the bottom of the slide.
    r = r & "top=" & Format$(shp.Top, "0.0") & vbLf
    r = r & "onslide_ok=" & Abs(CLng(shp.Top >= 0 And _
                                     shp.Top + shp.Height <= modSpec.SLIDE_H + 0.5)) & vbLf

    ' --- autocorrect repair ---------------------------------------------------
    ' Exactly what PowerPoint produces when you type quotes into a shape.
    shp.TextFrame.TextRange.text = "a = " & ChrW(&H201C) & "hi" & ChrW(&H201D) & _
                                   vbCr & "b = " & ChrW(&H2018) & "yo" & ChrW(&H2019)
    shp.Select
    modRibbon.DoStylize
    r = r & "curly_gone=" & Abs(CLng( _
            InStr(shp.TextFrame.TextRange.text, ChrW(&H201C)) = 0 And _
            InStr(shp.TextFrame.TextRange.text, ChrW(&H2018)) = 0)) & vbLf
    ' The repair only matters if the strings then colour AS strings.
    r = r & "string_mask=" & Quoted(Split(modLexer.MaskOf( _
            shp.TextFrame.TextRange.text, RunLang()), vbLf)(0)) & vbLf

    ' Put the real sample back for the render.
    shp.TextFrame.TextRange.text = modBlock.NormalizeParagraphs(code)
    shp.Select
    modRibbon.DoStylize

    ' AllShapes, not sld.Shapes: Stylize groups a block with its numbers and
    ' guides, so the block is no longer a top-level shape.
    For Each shp In modBlock.AllShapes(sld)
        If modBlock.IsCodeBlock(shp) Then blocks = blocks + 1
    Next shp
    r = r & "blocks_on_slide=" & blocks & vbLf

    sld.Export pngPath, "PNG", 1920, 1080
    RibbonSliceTest = r & "done=1"
    Exit Function

Failed:
    RibbonSliceTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Tags on every shape, so persistence across save/close/reopen can be checked
' from the outside.
Public Function ReportTags() As String
    Dim sld As Slide, shp As Shape, r As String
    On Error GoTo Failed
    For Each sld In Application.ActivePresentation.Slides
        For Each shp In modBlock.AllShapes(sld)
            r = r & sld.SlideIndex & " " & shp.Name & _
                    " block=" & shp.Tags(modBlock.TAG_BLOCK) & _
                    " lang=" & shp.Tags(modBlock.TAG_LANG) & _
                    " id=" & shp.Tags(modBlock.TAG_ID) & vbLf
        Next shp
    Next sld
    ReportTags = r
    Exit Function
Failed:
    ReportTags = "ERROR " & Err.Number & ": " & Err.Description
End Function

' 1 when the current selection resolves to a code block, 0 otherwise.
Private Function SelectionResolves() As Long
    Dim problem As String, found As Shape
    On Error Resume Next
    Set found = modBlock.SelectedBlock(problem)
    On Error GoTo 0
    SelectionResolves = Abs(CLng(Not found Is Nothing))
End Function

Private Function Quoted(ByVal s As String) As String
    Quoted = Chr$(34) & s & Chr$(34)
End Function

' Does autofit keep the shape's left and top anchored while it resizes, or does
' it grow around the centre? Determines whether a hugging block stays put.
Public Function AutofitProbe(ByVal dummy As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String

    On Error GoTo Failed
    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)

    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, 200, 200, 400, 100)
    shp.TextFrame.WordWrap = msoFalse
    shp.TextFrame.AutoSize = ppAutoSizeShapeToFitText
    shp.TextFrame.VerticalAnchor = msoAnchorTop
    r = "before  L=" & Format$(shp.Left, "0.0") & " T=" & Format$(shp.Top, "0.0") & _
        " W=" & Format$(shp.Width, "0.0") & " H=" & Format$(shp.Height, "0.0") & vbLf

    shp.TextFrame.TextRange.text = "x = 1" & vbCr & "y = 2" & vbCr & "print(a_much_longer_line)"
    modBlock.FormatBlockText shp, modSpec.BASE_SIZE
    r = r & "after3  L=" & Format$(shp.Left, "0.0") & " T=" & Format$(shp.Top, "0.0") & _
        " W=" & Format$(shp.Width, "0.0") & " H=" & Format$(shp.Height, "0.0") & vbLf

    shp.TextFrame.TextRange.text = shp.TextFrame.TextRange.text & vbCr & "z = 3" & vbCr & "w = 4"
    r = r & "after5  L=" & Format$(shp.Left, "0.0") & " T=" & Format$(shp.Top, "0.0") & _
        " W=" & Format$(shp.Width, "0.0") & " H=" & Format$(shp.Height, "0.0") & vbLf

    r = r & "spec_h_5_lines=" & Format$(modSpec.SpecHeight(modSpec.BASE_SIZE, 5), "0.0")
    AutofitProbe = r
    Exit Function
Failed:
    AutofitProbe = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Turns the gutter on, renders, and reports whether the numbers line up with
' the code. Alignment is the whole risk here, so it is measured rather than
' eyeballed: the gutter's own line boxes must step identically to the block's.
Public Function GutterTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, g As Shape
    Dim r As String, code As String, size As Single

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize
    r = "gutter_before=" & Abs(CLng(modGutter.HasGutter(shp))) & vbLf

    shp.Select
    modRibbon.DoToggleGutter
    r = r & "gutter_after=" & Abs(CLng(modGutter.HasGutter(shp))) & vbLf

    Set g = modGutter.FindGutter(shp)
    If g Is Nothing Then
        GutterTest = r & "no gutter shape"
        Exit Function
    End If

    size = modBlock.BlockFontSize(shp)
    r = r & "lines=" & modBlock.CountLines(shp.TextFrame.TextRange.text) & vbLf
    r = r & "gutter_lines=" & modBlock.CountLines(g.TextFrame.TextRange.text) & vbLf
    r = r & "top_match=" & Abs(CLng(Abs(g.Top - shp.Top) < 0.5)) & _
            " left_match=" & Abs(CLng(Abs(g.Left - shp.Left) < 0.5)) & vbLf
    r = r & "margins_match=" & Abs(CLng(Abs(g.TextFrame.MarginTop - shp.TextFrame.MarginTop) < 0.5)) & vbLf
    r = r & "code_margin_left=" & Format$(shp.TextFrame.MarginLeft, "0.0") & _
            " gutter_width=" & Format$(g.Width, "0.0") & vbLf
    r = r & "block_width=" & Format$(shp.Width, "0.0") & vbLf

    sld.Export pngPath, "PNG", 1920, 1080

    ' Toggling off must hand the space back.
    shp.Select
    modRibbon.DoToggleGutter
    r = r & "gutter_off=" & Abs(CLng(modGutter.HasGutter(shp))) & _
            " margin_restored=" & Format$(shp.TextFrame.MarginLeft, "0.0") & vbLf

    GutterTest = r
    Exit Function
Failed:
    GutterTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Drives the size ladder through the real commands, including both guards from
' PLAN.md section 5a: Larger capped at the fitting size, and Fit warning rather
' than silently shrinking below the teaching floor.
Public Function SizeTest(ByVal srcPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim r As String, code As String, i As Long, before As Single

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    r = "start=" & Format$(modBlock.BlockFontSize(shp), "0") & vbLf
    r = r & "lines=" & modBlock.CountLines(shp.TextFrame.TextRange.text) & _
            " cols=" & modBlock.LongestLine(shp.TextFrame.TextRange.text) & vbLf

    ' Smaller should walk down the ladder one rung at a time.
    modRibbon.DoSizeDown
    r = r & "after_smaller=" & Format$(modBlock.BlockFontSize(shp), "0") & vbLf

    ' Larger repeatedly, until the cap stops it. Ten steps is more rungs than
    ' the ladder has, so this must terminate at the fitting size.
    For i = 1 To 10
        modRibbon.DoSizeUp
    Next i
    r = r & "after_10x_larger=" & Format$(modBlock.BlockFontSize(shp), "0") & vbLf
    r = r & "cap_warned=" & Quoted(modRibbon.LastWarning()) & vbLf
    r = r & "fits_width=" & Abs(CLng(shp.Width <= modSpec.CONTENT_W + 0.5)) & _
            " fits_height=" & Abs(CLng(shp.Height <= modSpec.CONTENT_H + 0.5)) & vbLf

    ' Fit from a deliberately silly size.
    modBlock.ApplySize shp, 10
    shp.Select
    modRibbon.DoFit
    r = r & "fit_chose=" & Format$(modBlock.BlockFontSize(shp), "0") & vbLf
    r = r & "fit_warning=" & Quoted(Replace(modRibbon.LastWarning(), vbCrLf, " ")) & vbLf

    SizeTest = r
    Exit Function
Failed:
    SizeTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Writes a trace line after each step. A trace FILE survives a hang, which a
' return value does not - so this shows where execution stops rather than only
' that it stopped.
Public Function MiniTest(ByVal srcPath As String, ByVal tracePath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, code As String

    modRibbon.SetQuiet True
    Trace tracePath, "start"
    code = ReadTextFile(srcPath)
    Trace tracePath, "read " & Len(code)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    Trace tracePath, "slide added"

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    Trace tracePath, "block created"

    modRender.ApplyHighlight shp, "python"
    Trace tracePath, "highlighted"

    modBlock.ResizeToContent shp
    Trace tracePath, "resized"

    modGutter.SyncGutter shp
    Trace tracePath, "gutter synced (no-op)"

    modGuides.DrawGuides shp
    Trace tracePath, "guides drawn"

    modGutter.SyncGutter shp, True
    Trace tracePath, "gutter created"

    modGuides.DrawGuides shp
    Trace tracePath, "guides redrawn"

    MiniTest = "done"
End Function

Private Sub Trace(ByVal path As String, ByVal msg As String)
    Dim f As Integer
    f = FreeFile
    Open path For Append As #f
    Print #f, msg
    Close #f
End Sub

' The flow that actually breaks things: build a block, turn the gutter on,
' THEN edit the text and press Stylize again. Reported symptoms were a missing
' last number and misplaced guides, both of which only appear on the second
' pass, not on a block built in one go.
Public Function EditFlowTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, g As Shape
    Dim r As String, code As String, grp As Shape

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    ' Start from THREE lines only.
    Set shp = modBlock.CreateBlock(sld, "x = 1" & vbCr & "y = 2" & vbCr & "z = 3", _
                                   modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize
    shp.Select
    modRibbon.DoToggleGutter
    r = "start_lines=3 gutter=" & Abs(CLng(modGutter.HasGutter(shp))) & vbLf

    ' Now edit it to the full sample and Stylize again - the reported flow.
    ' The last break is a SOFT one (vertical tab), which is what Shift+Enter and
    ' pasted text produce, and what made the counts drift.
    modBlock.UngroupParts shp
    shp.TextFrame.TextRange.text = modBlock.NormalizeParagraphs(code) & _
                                   Chr$(11) & "print(""help"")"
    shp.Select
    modRibbon.DoStylize

    Set g = modGutter.FindGutter(shp)
    r = r & "code_lines=" & modBlock.CountLines(shp.TextFrame.TextRange.text) & vbLf
    If g Is Nothing Then
        r = r & "gutter LOST" & vbLf
    Else
        r = r & "gutter_numbers=" & modBlock.CountLines(g.TextFrame.TextRange.text) & vbLf
        r = r & "gutter_h=" & Format$(g.Height, "0.0") & _
                " needed=" & Format$(modSpec.SpecHeight(modBlock.BlockFontSize(shp), _
                     modBlock.CountLines(shp.TextFrame.TextRange.text)), "0.0") & vbLf
        r = r & "gutter_tall_enough=" & Abs(CLng(g.Height >= _
                modSpec.SpecHeight(modBlock.BlockFontSize(shp), _
                modBlock.CountLines(shp.TextFrame.TextRange.text)) - 0.5)) & vbLf
    End If

    ' Grouped?
    Dim k As Long
    r = r & "toplevel_shapes=" & sld.Shapes.count & vbLf
    For k = 1 To sld.Shapes.count
        r = r & "  [" & k & "] type=" & sld.Shapes(k).Type & _
                " name=" & sld.Shapes(k).Name
        If sld.Shapes(k).Type = msoGroup Then r = r & " items=" & sld.Shapes(k).GroupItems.count
        r = r & vbLf
    Next k
    Set grp = modBlock.ParentGroup(shp)
    r = r & "grouped=" & Abs(CLng(Not grp Is Nothing)) & _
            " why=" & Quoted(modBlock.LastGroupError) & vbLf
    If Not grp Is Nothing Then
        r = r & "group_items=" & grp.GroupItems.count & vbLf
        ' Selecting the GROUP must still find the block.
        grp.Select
        Dim problem As String, found As Shape
        Set found = modBlock.SelectedBlock(problem)
        r = r & "group_resolves=" & Abs(CLng(Not found Is Nothing)) & vbLf
    End If

    ' Report the geometry that has to line up, measured from PowerPoint's own
    ' layout rather than from the font metric.
    r = r & "code_first_top=" & Format$(shp.TextFrame.TextRange.Characters(1, 1).BoundTop, "0.00") & vbLf
    If Not g Is Nothing Then
        r = r & "num_first_top=" & Format$(g.TextFrame.TextRange.Characters(1, 1).BoundTop, "0.00") & vbLf
        r = r & "baseline_delta=" & Format$( _
                shp.TextFrame.TextRange.Characters(1, 1).BoundTop - _
                g.TextFrame.TextRange.Characters(1, 1).BoundTop, "0.00") & vbLf
    End If
    r = r & "code_origin_x=" & Format$(shp.TextFrame.TextRange.Characters(1, 1).BoundLeft, "0.00") & vbLf
    r = r & "spec_origin_x=" & Format$(shp.Left + shp.TextFrame.MarginLeft, "0.00") & vbLf

    ' Focus must survive a command, or every action costs a re-selection - and
    ' worse, the NEXT command silently does nothing because nothing is selected.
    r = r & "sel_after_highlight=" & SelectionResolves() & vbLf

    ' Two size steps WITHOUT re-selecting in between. This is the reported
    ' complaint: pressing Larger twice should not need a click between.
    modRibbon.DoSizeDown
    r = r & "size_after_1=" & Format$(modBlock.BlockFontSize(shp), "0") & vbLf
    modRibbon.DoSizeDown
    r = r & "size_after_2_no_reselect=" & Format$(modBlock.BlockFontSize(shp), "0") & vbLf
    r = r & "sel_after_resize=" & SelectionResolves() & vbLf

    sld.Export pngPath, "PNG", 1920, 1080

    ' Resizing must keep the colours.
    shp.Select
    modRibbon.DoSizeDown
    r = r & "after_smaller_size=" & Format$(modBlock.BlockFontSize(shp), "0") & vbLf
    r = r & "still_coloured=" & Abs(CLng(modLexer.MaskOf(shp.TextFrame.TextRange.text, _
            RunLang()) <> "")) & vbLf

    EditFlowTest = r
    Exit Function
Failed:
    EditFlowTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Hide a range, check the cover appears and keeps the layout, then Reveal and
' check the answer slide follows with nothing hidden.
Public Function HideTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, a As Long, b As Long, endStart As Long, endLen As Long
    Dim covers As Long, s2 As Shape, hBefore As Single

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize
    hBefore = shp.Height

    ' Hide lines 2 and 3.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, b
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 3, endStart, endLen
    shp.TextFrame.TextRange.Characters(a, (endStart + endLen) - a).Select
    modRibbon.DoHide

    r = "hidden_tag=" & Quoted(modBlock.GetHidden(shp)) & vbLf
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modBlock.TAG_COVER_OF) = shp.Tags(modBlock.TAG_ID) Then covers = covers + 1
    Next s2
    r = r & "cover_shapes=" & covers & " (one per run of hidden lines)" & vbLf
    r = r & "layout_unchanged=" & Abs(CLng(Abs(shp.Height - hBefore) < 0.5)) & vbLf

    ' The panel must cover every hidden line's INK, not its computed line box.
    ' Descenders hang below the box, and a two-point leak of code under an
    ' opaque panel is invisible here and plain on a projector.
    Dim cov As Shape, lt As Single, lh As Single, leak As Single, ln As Long
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modBlock.TAG_COVER_OF) = shp.Tags(modBlock.TAG_ID) Then Set cov = s2
    Next s2
    If Not cov Is Nothing Then
        For ln = 2 To 3
            If modRender.LineBounds(shp, ln, lt, lh) Then
                If cov.Top - lt > leak Then leak = cov.Top - lt
                If (lt + lh) - (cov.Top + cov.Height) > leak Then
                    leak = (lt + lh) - (cov.Top + cov.Height)
                End If
            End If
        Next ln
        r = r & "ink_leak_pt=" & Format$(leak, "0.00") & vbLf

        ' What the PREDICTED line box would have leaked, so the measurement is
        ' shown to be earning its keep rather than assumed to.
        Dim cSize As Single, cLine As Single, cPad As Single
        Dim cTop As Single, cBot As Single, wouldLeak As Single
        cSize = modBlock.BlockFontSize(shp)
        cLine = modSpec.SpecLine(cSize)
        cPad = modSpec.SpecPad(cSize)
        cTop = shp.Top + cPad + (2 - 1) * cLine
        cBot = shp.Top + cPad + 3 * cLine
        For ln = 2 To 3
            If modRender.LineBounds(shp, ln, lt, lh) Then
                If cTop - lt > wouldLeak Then wouldLeak = cTop - lt
                If (lt + lh) - cBot > wouldLeak Then wouldLeak = (lt + lh) - cBot
            End If
        Next ln
        r = r & "leak_if_predicted_pt=" & Format$(wouldLeak, "0.00") & vbLf
    End If

    sld.Export pngPath, "PNG", 1920, 1080

    ' Reveal adds the answer slide.
    shp.Select
    modRibbon.DoReveal
    r = r & "slides=" & pres.Slides.count & vbLf
    Dim target As Shape
    For Each s2 In modBlock.AllShapes(pres.Slides(pres.Slides.count))
        If s2.Tags(modBlock.TAG_BLOCK) = "1" Then Set target = s2
    Next s2
    If Not target Is Nothing Then
        r = r & "answer_slide_hidden=" & Quoted(modBlock.GetHidden(target)) & vbLf
    End If

    HideTest = r
    Exit Function
Failed:
    HideTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Attach two notes, then check the three things that can go wrong with them:
' they must not overlap, a note the user drags must stay where it was put, and
' a note must follow its line when the block is resized.
Public Function NoteTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, a As Long, b As Long, s2 As Shape
    Dim n1 As Shape, n2 As Shape, leaders As Long
    Dim draggedX As Single, draggedY As Single
    Dim size As Single, lineY As Single, beside As Boolean

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize

    ' Two notes on ADJACENT lines, which is the case that makes them collide.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 1, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNote
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNote

    r = "notes=" & modNote.NoteCount(shp) & vbLf
    Set n1 = modNote.FindNote(shp, 1)
    Set n2 = modNote.FindNote(shp, 2)
    If n1 Is Nothing Or n2 Is Nothing Then
        NoteTest = r & "ERROR: a note is missing"
        Exit Function
    End If

    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modNote.TAG_LEADER_OF) = shp.Tags(modBlock.TAG_ID) Then leaders = leaders + 1
    Next s2
    r = r & "leaders=" & leaders & vbLf
    r = r & "no_overlap=" & Abs(CLng(n2.Top >= n1.Top + n1.Height - 0.5)) & vbLf
    ' A block too wide to leave a margin puts its notes below itself instead,
    ' where they cannot be level with their line. Report which happened, so the
    ' level check below is read against the right expectation.
    beside = (n1.Left >= shp.Left + shp.Width) Or _
             (n1.Left + n1.Width <= shp.Left)
    r = r & "side=" & IIf(beside, "beside", "below") & vbLf

    sld.Export pngPath, "PNG", 1920, 1080

    ' The leader must be a real connector, attached at both ends, and it must
    ' stretch when the note moves - WITHOUT a Stylize. That is the whole reason
    ' for using a connector instead of a drawn line.
    '
    ' BOTH shapes are re-found immediately before measuring. An earlier version
    ' reused references from before the last regroup and measured nothing
    ' moving, which read as the connector failing when it was the test.
    ' Matched by what it is ATTACHED to, not just by tag. There are two notes
    ' here and therefore two leaders, and taking whichever came last measured
    ' the other note's connector against this note's movement.
    Dim lead As Shape, s3 As Shape
    Set n2 = modNote.FindNote(shp, 2)
    For Each s3 In modBlock.AllShapes(sld)
        If s3.Tags(modNote.TAG_LEADER_OF) = shp.Tags(modBlock.TAG_ID) Then
            On Error Resume Next
            If s3.ConnectorFormat.EndConnectedShape.Name = n2.Name Then Set lead = s3
            On Error GoTo Failed
        End If
    Next s3

    r = r & "leader_connected=" & _
            Abs(CLng(lead.ConnectorFormat.BeginConnected = msoTrue And _
                     lead.ConnectorFormat.EndConnected = msoTrue)) & vbLf

    Dim wasW As Single, wasH As Single
    wasW = lead.Width: wasH = lead.Height
    n2.Left = n2.Left + 60
    n2.Top = n2.Top + 90
    ' The box has to change; ConnectorProbe pins down that it changes by exactly
    ' the distance moved. Here the routing may pick a different pair of sites,
    ' since this note has another one stacked above it.
    r = r & "leader_stretched=" & _
            Abs(CLng(Abs(lead.Width - wasW) > 2 Or Abs(lead.Height - wasH) > 2)) & vbLf
    n2.Left = n2.Left - 60
    n2.Top = n2.Top - 90

    ' Deleting a note takes its connector and its anchor with it.
    Dim before As Long, after As Long
    before = modNote.NoteCount(shp)
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoDeleteNote
    after = modNote.NoteCount(shp)
    Dim leftovers As Long
    For Each s3 In modBlock.AllShapes(sld)
        If s3.Tags(modNote.TAG_LEADER_OF) = shp.Tags(modBlock.TAG_ID) Then leftovers = leftovers + 1
        If s3.Tags(modNote.TAG_ANCHOR_OF) = shp.Tags(modBlock.TAG_ID) Then leftovers = leftovers + 1
    Next s3
    r = r & "deleted_one=" & Abs(CLng(after = before - 1)) & vbLf
    ' One note left, so one connector and one anchor.
    r = r & "no_orphan_parts=" & Abs(CLng(leftovers = after * 2)) & vbLf

    ' Remake it, so the rest of the test has two notes again.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNote
    Set n2 = modNote.FindNote(shp, 2)

    ' A dragged note must survive the next Stylize.
    ' Downward, the way someone actually drags a note clear of the one above it.
    n2.Left = n2.Left + 30
    n2.Top = n2.Top + 70
    draggedX = n2.Left
    draggedY = n2.Top
    shp.Select
    modRibbon.DoStylize
    Set n2 = modNote.FindNote(shp, 2)
    r = r & "drag_kept=" & Abs(CLng(Abs(n2.Left - draggedX) < 2 And _
                                    Abs(n2.Top - draggedY) < 2)) & vbLf

    ' And an undragged note must stay LEVEL WITH ITS LINE when the size changes,
    ' which is the property that matters - not merely that it moved.
    shp.Select
    modRibbon.DoSizeDown
    Set n1 = modNote.FindNote(shp, 1)
    size = modBlock.BlockFontSize(shp)
    lineY = shp.Top + modSpec.SpecPad(size) + (1 - 0.5) * modSpec.SpecLine(size)
    r = r & "size_after=" & Format$(size, "0") & vbLf
    If beside Then
        r = r & "level_with_line=" & _
                Abs(CLng(Abs((n1.Top + n1.Height / 2) - lineY) < 2)) & vbLf
    Else
        r = r & "level_with_line=n/a (placed below the block)" & vbLf
    End If

    ' On a walkthrough slide the block itself is enough: the line comes from the
    ' emphasis, and from its LAST line, which is where Build up has got to.
    modBlock.SetEmphasis shp, "3,4,5"
    shp.Select
    modRibbon.DoStylize
    ' The SHAPE, not text inside it - that is the gesture being tested.
    Dim grp As Shape
    Set grp = modBlock.ParentGroup(shp)
    If grp Is Nothing Then shp.Select Else grp.Select
    modRibbon.DoNote
    r = r & "note_from_emphasis=" & _
            Abs(CLng(Not modNote.FindNote(shp, 5) Is Nothing)) & vbLf

    ' Colour and size reach a note that has been SINGLED OUT - here by putting
    ' the cursor on its line, which is the gesture that works on a note inside
    ' a group. With nothing singled out they would correctly do nothing.
    modBlock.SetEmphasis shp, ""
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 1, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNoteColor 5                      ' Paper, the one light preset
    Set n1 = modNote.FindNote(shp, 1)
    r = r & "colour_applied=" & _
            Abs(CLng(n1.fill.ForeColor.RGB = ThemeNotePreset(5))) & vbLf
    ' A light fill must get dark words, or the note is unreadable.
    r = r & "text_follows_fill=" & _
            Abs(CLng(n1.TextFrame.TextRange.Font.Color.RGB = ThemeTextOn(ThemeNotePreset(5)))) & vbLf

    modBlock.LineCharRange shp.TextFrame.TextRange.text, 1, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNoteSize 32
    Set n1 = modNote.FindNote(shp, 1)
    r = r & "size_applied=" & Format$(n1.TextFrame.TextRange.Font.size, "0") & vbLf

    sld.Export Replace(pngPath, ".png", "-styled.png"), "PNG", 1920, 1080

    ' Strip must not eat the words.
    shp.Select
    modRibbon.DoStrip
    r = r & "notes_after_strip=" & modNote.NoteCount(shp) & vbLf

    NoteTest = r
    Exit Function
Failed:
    NoteTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' EVERYTHING on one block at once, then Stylize twice.
'
' The parts have all been tested apart. What has not is the interaction: eight
' kinds of child shape - gutter, guides, band, cover, note, leader, anchor,
' arrow - competing for one group and one z-order, on a pipeline that ungroups
' and regroups on every pass. Two things can go wrong there and neither shows up
' in a single-feature test. A part can be LOST, because GroupParts gathers by
' tag and a tag it does not know about is left behind on the slide. And a part
' can be DUPLICATED, because anything created rather than repositioned piles up.
'
' So it counts every kind, styles again, and counts again. Equal counts across a
' second pass is the property that matters: the pipeline has to be idempotent,
' since Stylize is the button you press constantly.
Public Function EverythingTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, a As Long, b As Long, e1 As Long, e2 As Long

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Left = 90
    shp.Select
    modRibbon.DoStylize

    ' Numbers and guides.
    shp.Select
    modRibbon.DoToggleGutter
    shp.Select
    modRibbon.DoToggleGuides

    ' Emphasis, in bold.
    modOptions.SetEmphasisBold True
    modBlock.SetEmphasis shp, "4,5"

    ' A hidden run.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 8, a, b
    shp.TextFrame.TextRange.Characters(a, b).Select
    modRibbon.DoHide

    ' Two notes and two arrows, on different lines.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 1, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNote
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 6, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNote
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 3, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoArrow
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 10, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoArrow

    shp.Select
    modRibbon.DoStylize
    r = "after_first=" & PartCensus(shp) & vbLf
    r = r & "all_in_one_group=" & Abs(CLng(LooseParts(shp) = 0)) & vbLf
    e1 = TotalParts(shp)

    sld.Export pngPath, "PNG", 1920, 1080

    ' Again. Nothing may appear and nothing may vanish.
    shp.Select
    modRibbon.DoStylize
    r = r & "after_second=" & PartCensus(shp) & vbLf
    e2 = TotalParts(shp)
    r = r & "idempotent=" & Abs(CLng(e1 = e2)) & vbLf
    r = r & "still_one_group=" & Abs(CLng(LooseParts(shp) = 0)) & vbLf

    modOptions.SetEmphasisBold False
    EverythingTest = r
    Exit Function
Failed:
    EverythingTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' One line naming every kind of part and how many there are.
Private Function PartCensus(ByVal shp As Shape) As String
    Dim id As String, sld As Slide, s2 As Shape
    Dim gut As Long, gui As Long, bnd As Long, cov As Long
    Dim nte As Long, led As Long, anc As Long, arw As Long

    id = shp.Tags(modBlock.TAG_ID)
    Set sld = modGutter.OwningSlide(shp)
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modGutter.TAG_GUTTER_OF) = id Then gut = gut + 1
        If s2.Tags(modGuides.TAG_GUIDE_OF) = id Then gui = gui + 1
        If s2.Tags(modBlock.TAG_BAND_OF) = id Then bnd = bnd + 1
        If s2.Tags(modBlock.TAG_COVER_OF) = id Then cov = cov + 1
        If s2.Tags(modNote.TAG_NOTE_OF) = id Then nte = nte + 1
        If s2.Tags(modNote.TAG_LEADER_OF) = id Then led = led + 1
        If s2.Tags(modNote.TAG_ANCHOR_OF) = id Then anc = anc + 1
        If s2.Tags(modArrow.TAG_ARROW_OF) = id Then arw = arw + 1
    Next s2

    PartCensus = "gutter=" & gut & " guides=" & gui & " bands=" & bnd & _
                 " covers=" & cov & " notes=" & nte & " leaders=" & led & _
                 " anchors=" & anc & " arrows=" & arw
End Function

Private Function TotalParts(ByVal shp As Shape) As Long
    Dim id As String, sld As Slide, s2 As Shape, n As Long
    id = shp.Tags(modBlock.TAG_ID)
    Set sld = modGutter.OwningSlide(shp)
    For Each s2 In modBlock.AllShapes(sld)
        If PartOf(s2, id) Then n = n + 1
    Next s2
    TotalParts = n
End Function

' Parts sitting on the SLIDE rather than inside the block's group. Any of these
' is a tag GroupParts does not know about, and it means the part stops moving
' with the block.
Private Function LooseParts(ByVal shp As Shape) As Long
    Dim id As String, sld As Slide, i As Long, n As Long
    id = shp.Tags(modBlock.TAG_ID)
    Set sld = modGutter.OwningSlide(shp)
    For i = 1 To sld.Shapes.count
        If PartOf(sld.Shapes(i), id) Then n = n + 1
    Next i
    LooseParts = n
End Function

Private Function PartOf(ByVal s2 As Shape, ByVal id As String) As Boolean
    PartOf = (s2.Tags(modGutter.TAG_GUTTER_OF) = id) Or _
             (s2.Tags(modGuides.TAG_GUIDE_OF) = id) Or _
             (s2.Tags(modBlock.TAG_BAND_OF) = id) Or _
             (s2.Tags(modBlock.TAG_COVER_OF) = id) Or _
             (s2.Tags(modNote.TAG_NOTE_OF) = id) Or _
             (s2.Tags(modNote.TAG_LEADER_OF) = id) Or _
             (s2.Tags(modNote.TAG_ANCHOR_OF) = id) Or _
             (s2.Tags(modArrow.TAG_ARROW_OF) = id)
End Function

' Line numbering that starts somewhere other than 1, for code split across
' slides. The width case is the one worth asserting: a block starting at 98
' needs three digits where its line COUNT would only ask for two.
Public Function FirstLineTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, g As Shape, w1 As Single, w2 As Single, nums As String

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize
    shp.Select
    modRibbon.DoToggleGutter

    Set g = modGutter.FindGutter(shp)
    w1 = g.Width
    r = "default_first=" & modGutter.FirstLine(shp) & vbLf
    r = r & "numbers_from_1=" & Quoted(Left$(modGutter.NumberColumn( _
            shp.TextFrame.TextRange.text), 3)) & vbLf

    ' Start at 98, so the numbers cross into three digits partway down.
    shp.Select
    modRibbon.DoFirstLine 98
    Set g = modGutter.FindGutter(shp)
    w2 = g.Width
    nums = modGutter.NumberColumn(shp.TextFrame.TextRange.text, modGutter.FirstLine(shp))

    r = r & "first_after=" & modGutter.FirstLine(shp) & vbLf
    r = r & "starts_at_98=" & Abs(CLng(Left$(nums, 2) = "98")) & vbLf
    r = r & "gutter_widened=" & Abs(CLng(w2 > w1)) & _
            " (" & Format$(w1, "0.0") & " -> " & Format$(w2, "0.0") & ")" & vbLf

    ' It has to survive a Stylize, like every other per-block choice.
    shp.Select
    modRibbon.DoStylize
    r = r & "survives_stylize=" & Abs(CLng(modGutter.FirstLine(shp) = 98)) & vbLf

    sld.Export pngPath, "PNG", 1920, 1080
    FirstLineTest = r
    Exit Function
Failed:
    FirstLineTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Output notes: the dress, the mark, and the hanging indent.
'
' The indent is the one worth MEASURING. Output long enough to wrap must
' continue flush with the first character after the arrow, not under the arrow -
' otherwise the second line reads as a second result. BoundLeft says where the
' characters actually landed, which is the only way to know.
Public Function OutputNoteTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim a As Long, b As Long, n As Shape, long_ As String
    Dim afterMark As Single, wrapped As Single, i As Long, ln2 As Long

    On Error GoTo Failed
    modRibbon.SetQuiet True

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, _
        "print(set(range(5)))" & vbCr & "print(sorted(words))", _
        modSpec.BASE_SIZE, "python")
    shp.Left = 60
    shp.Top = 140
    shp.Select
    modRibbon.DoStylize

    ' A short one: must not wrap.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 1, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoOutputNote
    Set n = modNote.FindNote(shp, 1)
    n.TextFrame.TextRange.text = modNote.OutputMark() & "{0, 1, 2, 3, 4}"
    modNote.StyleOutputNote n, modBlock.BlockFontSize(shp), 400

    r = "is_output=" & Abs(CLng(modNote.IsOutputNote(n))) & vbLf
    r = r & "mark_present=" & Abs(CLng(Left$(n.TextFrame.TextRange.text, 1) = ChrW(&H2192))) & vbLf
    r = r & "mark_is_green=" & _
            Abs(CLng(n.TextFrame.TextRange.Characters(1, 1).Font.Color.RGB = ThemeOutputMark())) & vbLf
    r = r & "terminal_fill=" & Abs(CLng(n.fill.ForeColor.RGB = ThemeOutputFill())) & vbLf
    r = r & "code_font=" & Abs(CLng(n.TextFrame.TextRange.Font.Name = THEME_FONT)) & vbLf
    r = r & "short_does_not_wrap=" & Abs(CLng(n.TextFrame.WordWrap = msoFalse)) & vbLf

    ' The mark comes back if it is typed over.
    n.TextFrame.TextRange.text = "{0, 1, 2, 3, 4}"
    modNote.StyleOutputNote n, modBlock.BlockFontSize(shp), 400
    r = r & "mark_restored=" & _
            Abs(CLng(Left$(n.TextFrame.TextRange.text, 1) = ChrW(&H2192))) & vbLf

    ' A long one: must wrap, and the wrapped line must start flush with the
    ' first character after the mark.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoOutputNote
    Set n = modNote.FindNote(shp, 2)
    long_ = "['alpha', 'bravo', 'charlie', 'delta', 'echo', 'foxtrot', 'golf']"
    n.TextFrame.TextRange.text = modNote.OutputMark() & long_
    modNote.StyleOutputNote n, modBlock.BlockFontSize(shp), 300

    r = r & "long_wraps=" & Abs(CLng(n.TextFrame.WordWrap = msoTrue)) & vbLf

    ' Character 3 is the first after "arrow + space". Then find the first
    ' character that PowerPoint has laid out on a lower line.
    afterMark = n.TextFrame.TextRange.Characters(3, 1).BoundLeft
    For i = 4 To Len(n.TextFrame.TextRange.text)
        If n.TextFrame.TextRange.Characters(i, 1).BoundTop > _
           n.TextFrame.TextRange.Characters(3, 1).BoundTop + 2 Then
            wrapped = n.TextFrame.TextRange.Characters(i, 1).BoundLeft
            ln2 = i
            Exit For
        End If
    Next i

    If ln2 = 0 Then
        r = r & "hanging_indent=n/a (it did not actually wrap)" & vbLf
    Else
        r = r & "first_after_mark_x=" & Format$(afterMark, "0.0") & _
                " wrapped_line_x=" & Format$(wrapped, "0.0") & vbLf
        r = r & "hanging_indent=" & Abs(CLng(Abs(wrapped - afterMark) < 2)) & vbLf
    End If

    ' Survives a Stylize, and the note colour control leaves it alone.
    shp.Select
    modRibbon.DoStylize
    Set n = modNote.FindNote(shp, 2)
    r = r & "survives_stylize=" & _
            Abs(CLng(modNote.IsOutputNote(n) And n.fill.ForeColor.RGB = ThemeOutputFill())) & vbLf

    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNoteColor 6                                  ' Crimson
    Set n = modNote.FindNote(shp, 2)
    r = r & "note_color_skips_output=" & _
            Abs(CLng(n.fill.ForeColor.RGB = ThemeOutputFill())) & vbLf

    ' Two output notes must stack, not pile up. Dressing them after placement
    ' made the placement record a stale size, and the next pass read that as a
    ' drag - which put both in the same spot.
    Dim n1 As Shape
    Set n1 = modNote.FindNote(shp, 1)
    Set n = modNote.FindNote(shp, 2)
    r = r & "two_outputs_stack=" & _
            Abs(CLng(n.Top >= n1.Top + n1.Height - 0.5)) & vbLf

    sld.Export pngPath, "PNG", 1920, 1080
    OutputNoteTest = r
    Exit Function
Failed:
    OutputNoteTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Round two on labelling an output note: the word tightened to the top, and
' four inline icons. An icon on the output's own line costs NO vertical space
' at all, where even a tight word costs a line.
Public Function OutputIconProbe(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim v As Long, r As String
    For v = 1 To 5
        r = r & v & "=" & OneIcon(v, pngPath) & vbLf
    Next v
    OutputIconProbe = r
End Function

Private Function OneIcon(ByVal v As Long, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim n1 As Shape, n2 As Shape, png As String, g As String

    On Error GoTo Failed
    modRibbon.SetQuiet True

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, _
        "print(set(range(5)))" & vbCr & "print(set(" & Chr$(34) & "hello" & Chr$(34) & "))", _
        modSpec.BASE_SIZE, "python")
    shp.Left = 120
    shp.Top = 150
    shp.Select
    modRibbon.DoStylize
    modBlock.UngroupParts shp

    Set n1 = OutputNote(sld, shp, 1, "{0, 1, 2, 3, 4}")
    Set n2 = OutputNote(sld, shp, 2, "{'h', 'e', 'l', 'o'}")

    Select Case v
        Case 1                                  ' the word, pulled to the top
            TightLabel n1, "output"
            TightLabel n2, "output"
        Case 2: g = ChrW(&H2192)                ' arrow
        Case 3: g = ChrW(&H21D2)                ' double arrow, "yields"
        Case 4: g = ChrW(&H25B6)                ' solid triangle
        Case Else: g = ChrW(&HBB)               ' guillemets
    End Select

    If Len(g) > 0 Then
        IconPrefix n1, g
        IconPrefix n2, g
    End If

    modNote.PlaceNotes shp

    png = Replace(pngPath, ".png", "-" & v & ".png")
    sld.Export png, "PNG", 1920, 1080
    OneIcon = "ok"
    Exit Function
Failed:
    OneIcon = "ERROR " & Err.Number & ": " & Err.Description
End Function

' The word as the note's first paragraph, pulled tight to the top.
'
' Two things make the space above it: the frame's top margin, and the label
' getting a full line of the OUTPUT's leading. Both have to shrink, or trimming
' one just moves the gap.
Private Sub TightLabel(ByVal note As Shape, ByVal caption As String)
    Dim sz As Single, before As String
    sz = LabelSize(note)
    before = note.TextFrame.TextRange.text

    note.TextFrame.MarginTop = 2
    note.TextFrame.TextRange.text = caption & vbCr & before
    With note.TextFrame.TextRange
        .Font.Name = THEME_FONT
        .Font.Color.RGB = RGB(204, 204, 204)
        With .Paragraphs(1)
            .Font.size = sz
            .Font.Color.RGB = RGB(130, 130, 130)
            With .ParagraphFormat
                .LineRuleWithin = msoFalse
                .SpaceWithin = sz * 1.05
                .LineRuleAfter = msoFalse
                .SpaceAfter = 0
            End With
        End With
    End With
End Sub

' A glyph on the output's own line, dim, before the value.
Private Sub IconPrefix(ByVal note As Shape, ByVal glyph As String)
    Dim before As String
    before = note.TextFrame.TextRange.text
    note.TextFrame.TextRange.text = glyph & " " & before
    With note.TextFrame.TextRange
        .Font.Name = THEME_FONT
        .Font.Color.RGB = RGB(204, 204, 204)
        .Characters(1, 1).Font.Color.RGB = RGB(135, 199, 107)
    End With
End Sub

' MOCKUPS for labelling an output note. Variant I won; this is only about how
' the note announces itself. Two notes, so a label that collides with the note
' above it shows up here rather than on a real slide.
Public Function OutputLabelProbe(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim v As Long, r As String
    For v = 1 To 5
        r = r & v & "=" & OneLabel(v, pngPath) & vbLf
    Next v
    OutputLabelProbe = r
End Function

Private Function OneLabel(ByVal v As Long, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim n1 As Shape, n2 As Shape, png As String

    On Error GoTo Failed
    modRibbon.SetQuiet True

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, _
        "print(set(range(5)))" & vbCr & "print(set(" & Chr$(34) & "hello" & Chr$(34) & "))", _
        modSpec.BASE_SIZE, "python")
    shp.Left = 120
    shp.Top = 150
    shp.Select
    modRibbon.DoStylize
    modBlock.UngroupParts shp

    Set n1 = OutputNote(sld, shp, 1, "{0, 1, 2, 3, 4}")
    Set n2 = OutputNote(sld, shp, 2, "{'h', 'e', 'l', 'o'}")
    modNote.PlaceNotes shp

    Select Case v
        Case 1                                  ' no label at all, the control
        Case 2                                  ' the word, above-left
            OutLabel sld, n1, "output", -1
            OutLabel sld, n2, "output", -1
        Case 3                                  ' the word, below-left
            OutLabel sld, n1, "output", 1
            OutLabel sld, n2, "output", 1
        Case 4                                  ' Jupyter's Out:, left of centre
            SideLabel sld, n1, "Out:"
            SideLabel sld, n2, "Out:"
        Case Else                               ' inside the note, top-left
            InsideLabel n1, "output"
            InsideLabel n2, "output"
    End Select

    png = Replace(pngPath, ".png", "-" & v & ".png")
    sld.Export png, "PNG", 1920, 1080
    OneLabel = "ok"
    Exit Function
Failed:
    OneLabel = "ERROR " & Err.Number & ": " & Err.Description
End Function

' side = -1 above the note, +1 below it. Left edges aligned with the note.
Private Sub OutLabel(ByVal sld As Slide, ByVal note As Shape, _
                     ByVal caption As String, ByVal side As Long)
    Dim t As Shape, sz As Single, h As Single
    sz = LabelSize(note)
    h = sz * 1.5

    Set t = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, note.Left, 0, 120, h)
    If side < 0 Then
        t.Top = note.Top - h + sz * 0.25
    Else
        t.Top = note.Top + note.Height - sz * 0.25
    End If
    StyleLabel t, caption, sz, ppAlignLeft
End Sub

' Vertically centred to the LEFT of the note, the way Jupyter puts Out[3]:.
Private Sub SideLabel(ByVal sld As Slide, ByVal note As Shape, ByVal caption As String)
    Dim t As Shape, sz As Single, w As Single
    sz = LabelSize(note)
    w = 60
    Set t = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                  note.Left - w - 6, note.Top + note.Height / 2 - sz, _
                                  w, sz * 2)
    StyleLabel t, caption, sz, ppAlignRight
End Sub

' Prefixed inside the note itself, dimmer than the output.
Private Sub InsideLabel(ByVal note As Shape, ByVal caption As String)
    Dim before As String
    before = note.TextFrame.TextRange.text
    note.TextFrame.TextRange.text = caption & vbCr & before
    With note.TextFrame.TextRange
        .Font.Name = THEME_FONT
        .Font.Color.RGB = RGB(204, 204, 204)
        .Characters(1, Len(caption)).Font.size = LabelSize(note)
        .Characters(1, Len(caption)).Font.Color.RGB = RGB(130, 130, 130)
    End With
End Sub

Private Function LabelSize(ByVal note As Shape) As Single
    Dim sz As Single
    sz = Round(note.TextFrame.TextRange.Font.size * 0.55, 0)
    If sz < 11 Then sz = 11
    LabelSize = sz
End Function

Private Sub StyleLabel(ByVal t As Shape, ByVal caption As String, _
                       ByVal sz As Single, ByVal align As Long)
    t.Line.Visible = msoFalse
    t.fill.Visible = msoFalse
    With t.TextFrame
        .WordWrap = msoFalse
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        With .TextRange
            .text = caption
            .Font.Name = THEME_FONT
            .Font.size = sz
            .Font.Color.RGB = RGB(130, 130, 130)
            .ParagraphFormat.Alignment = align
        End With
    End With
End Sub

' MOCKUPS for showing a program's output beside its code. Renders one slide
' per candidate so they can be compared as a set - judging one at a time is how
' you end up with six shades of slate.
'
' NOTHING here is product code. Every variant paints runs and drops shapes on
' top by hand, so no tag and no pipeline stage has to exist yet.
'
' The transcript is built inline rather than read from tests/samples, because it
' is a REPL transcript rather than a program and would fail the lexer mask diff
' if it entered the corpus.
Public Function OutputProbe(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim v As Long, r As String
    For v = 1 To 9
        r = r & Chr$(64 + v) & "=" & OneVariant(v, pngPath) & vbLf
    Next v
    OutputProbe = r
End Function

Private Function TranscriptText() As String
    TranscriptText = "print(set(range(5)))" & vbCr & _
                     "{0, 1, 2, 3, 4}" & vbCr & _
                     "print(set(" & Chr$(34) & "hello" & Chr$(34) & "))" & vbCr & _
                     "{'h', 'e', 'l', 'o'}"
End Function

Private Function OneVariant(ByVal v As Long, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, shp2 As Shape
    Dim png As String, plain As Long, i As Long

    On Error GoTo Failed
    modRibbon.SetQuiet True

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, TranscriptText(), modSpec.BASE_SIZE, "python")
    shp.Left = 120
    shp.Top = 120
    shp.Select
    modRibbon.DoStylize
    modBlock.UngroupParts shp

    plain = ThemeColor(tkDefault)

    Select Case v
        Case 1                                  ' A - plain, no band
            PaintLine shp, 2, plain
            PaintLine shp, 4, plain

        Case 2                                  ' B - plain on a LIGHTER band
            AddLineBand sld, shp, 2, 2, ThemeCoverColor()
            AddLineBand sld, shp, 4, 4, ThemeCoverColor()
            PaintLine shp, 2, plain
            PaintLine shp, 4, plain
            shp.ZOrder msoSendToBack

        Case 3                                  ' C - plain, thin rule in margin
            PaintLine shp, 2, plain
            PaintLine shp, 4, plain
            AddMarginRule sld, shp, 2
            AddMarginRule sld, shp, 4

        Case 4                                  ' D - dimmed, no band
            PaintLine shp, 2, ThemeDimmed(plain)
            PaintLine shp, 4, ThemeDimmed(plain)

        Case 5                                  ' E - a second, Stripped block
            shp.TextFrame.TextRange.text = "print(set(range(5)))" & vbCr & _
                                           "print(set(" & Chr$(34) & "hello" & Chr$(34) & "))"
            modBlock.FormatBlockText shp, modSpec.BASE_SIZE
            shp.Select
            modRibbon.DoStylize
            modBlock.UngroupParts shp
            Set shp2 = modBlock.CreateBlock(sld, _
                "{0, 1, 2, 3, 4}" & vbCr & "{'h', 'e', 'l', 'o'}", _
                modSpec.BASE_SIZE, "python")
            shp2.TextFrame.TextRange.Font.Color.RGB = plain
            shp2.Left = shp.Left
            shp2.Top = shp.Top + shp.Height + 14
            shp2.Width = shp.Width

        Case 6                                  ' F - the VS Code terminal
            ' Text set bright first: the panel goes OVER it and takes it back
            ' down. Nothing can be placed between a shape's fill and its own
            ' text, so a panel that darkens the background darkens the words
            ' with it - the same constraint the emphasis bands live under.
            PaintLine shp, 2, RGB(255, 255, 255)
            PaintLine shp, 4, RGB(255, 255, 255)
            TerminalPanel sld, shp, 2, 0.62, RGB(60, 60, 60), 1
            TerminalPanel sld, shp, 4, 0.62, RGB(60, 60, 60), 1

        Case 7                                  ' G - F, separation widened
            PaintLine shp, 2, RGB(255, 255, 255)
            PaintLine shp, 4, RGB(255, 255, 255)
            TerminalPanel sld, shp, 4, 0.34, RGB(110, 110, 110), 2
            TerminalPanel sld, shp, 2, 0.34, RGB(110, 110, 110), 2

        Case 8                                  ' H - all terminal, prompt column
            shp.fill.ForeColor.RGB = RGB(24, 24, 24)
            PaintLine shp, 2, RGB(204, 204, 204)
            PaintLine shp, 4, RGB(204, 204, 204)
            PromptColumn sld, shp

        Case Else                               ' I - notes carrying the output
            ' The text has to shrink to the two code lines FIRST. Notes are
            ' anchored to a line number, and anchoring to line 3 of a block
            ' about to become two lines long points at nothing.
            shp.TextFrame.TextRange.text = "print(set(range(5)))" & vbCr & _
                                           "print(set(" & Chr$(34) & "hello" & Chr$(34) & "))"
            modBlock.FormatBlockText shp, modSpec.BASE_SIZE
            modRender.ApplyHighlight shp, "python"
            modBlock.ResizeToContent shp
            OutputNote sld, shp, 1, "{0, 1, 2, 3, 4}"
            OutputNote sld, shp, 2, "{'h', 'e', 'l', 'o'}"
            modNote.PlaceNotes shp
    End Select

    png = Replace(pngPath, ".png", "-" & Chr$(64 + v) & ".png")
    sld.Export png, "PNG", 1920, 1080
    OneVariant = "ok"
    Exit Function
Failed:
    OneVariant = "ERROR " & Err.Number & ": " & Err.Description
End Function

'--- helpers, mockup only -----------------------------------------------------

Private Sub PaintLine(ByVal shp As Shape, ByVal ln As Long, ByVal rgbColor As Long)
    Dim a As Long, l As Long
    modBlock.LineCharRange shp.TextFrame.TextRange.text, ln, a, l
    If a >= 1 And l >= 1 Then
        shp.TextFrame.TextRange.Characters(a, l).Font.Color.RGB = rgbColor
    End If
End Sub

' Full-width band behind a run of lines, sized from the MEASURED ink the way the
' cover panels are.
Private Function AddLineBand(ByVal sld As Slide, ByVal shp As Shape, _
                             ByVal ln1 As Long, ByVal ln2 As Long, _
                             ByVal rgbColor As Long) As Shape
    Dim t1 As Single, h1 As Single, t2 As Single, h2 As Single
    Dim size As Single, lineH As Single, pad As Single, y0 As Single, y1 As Single
    Dim b As Shape

    size = modBlock.BlockFontSize(shp)
    lineH = modSpec.SpecLine(size)
    pad = modSpec.SpecPad(size)

    y0 = shp.Top + pad + (ln1 - 1) * lineH
    y1 = shp.Top + pad + ln2 * lineH
    If modRender.LineBounds(shp, ln1, t1, h1) Then If t1 < y0 Then y0 = t1
    If modRender.LineBounds(shp, ln2, t2, h2) Then If t2 + h2 > y1 Then y1 = t2 + h2

    Set b = sld.Shapes.AddShape(msoShapeRectangle, shp.Left, y0, shp.Width, y1 - y0)
    With b
        .fill.Solid
        .fill.ForeColor.RGB = rgbColor
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
    End With
    Set AddLineBand = b
End Function

Private Sub AddMarginRule(ByVal sld As Slide, ByVal shp As Shape, ByVal ln As Long)
    Dim t As Single, h As Single, ru As Shape, x As Single
    If Not modRender.LineBounds(shp, ln, t, h) Then Exit Sub
    x = shp.Left + modSpec.SpecPad(modBlock.BlockFontSize(shp)) * 0.5
    Set ru = sld.Shapes.AddLine(x, t, x, t + h)
    ru.Line.ForeColor.RGB = RGB(140, 140, 140)
    ru.Line.Weight = 2.5
End Sub

' Everything from the first output line down is darkened, with a border along
' its top - the arrangement VS Code uses for the terminal below the editor,
' where the panel is DARKER than the editor rather than lighter.
'
' Darkened by a translucent black overlay rather than an opaque panel, because
' an opaque one would have to sit either behind the block's own fill (invisible)
' or over its text (illegible). Transparency is the only layer available
' between the two - see modRender.ApplyEmphasis, which hit this first.
Private Sub TerminalPanel(ByVal sld As Slide, ByVal shp As Shape, _
                          ByVal outLine As Long, _
                          ByVal transp As Single, ByVal borderColor As Long, _
                          ByVal borderWeight As Single)
    Dim t As Single, h As Single, y0 As Single, y1 As Single
    Dim p As Shape, ln As Shape, lineH As Single, pad As Single

    If Not modRender.LineBounds(shp, outLine, t, h) Then Exit Sub
    lineH = modSpec.SpecLine(modBlock.BlockFontSize(shp))
    pad = modSpec.SpecPad(modBlock.BlockFontSize(shp))

    ' ONE PANEL PER OUTPUT RUN, not one from here down. VS Code's terminal has a
    ' single top edge because the editor is entirely above it - a transcript
    ' interleaves, so a single panel swallows the code that follows. This is the
    ' shape DrawCovers already uses.
    y0 = shp.Top + pad + (outLine - 1) * lineH
    y1 = shp.Top + pad + outLine * lineH
    If t < y0 Then y0 = t
    If t + h > y1 Then y1 = t + h

    Set p = sld.Shapes.AddShape(msoShapeRectangle, shp.Left, y0, shp.Width, y1 - y0)
    With p
        .fill.Solid
        .fill.ForeColor.RGB = RGB(0, 0, 0)
        .fill.Transparency = transp
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .ZOrder msoBringToFront
    End With
    Set ln = sld.Shapes.AddLine(shp.Left, y0, shp.Left + shp.Width, y0)
    ln.Line.ForeColor.RGB = borderColor
    ln.Line.Weight = borderWeight
    ln.ZOrder msoBringToFront
End Sub

' ">>>" beside the code lines and nothing beside the output, drawn in its own
' shape the way the line numbers are - so the block's text stays pure source.
Private Sub PromptColumn(ByVal sld As Slide, ByVal shp As Shape)
    Dim g As Shape, size As Single, pad As Single, w As Single
    size = modBlock.BlockFontSize(shp)
    pad = modSpec.SpecPad(size)
    w = 4 * size * 0.55

    shp.TextFrame.MarginLeft = pad + w
    Set g = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                  shp.Left, shp.Top, pad + w, shp.Height)
    g.Line.Visible = msoFalse
    g.fill.Visible = msoFalse
    With g.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeNone
        .VerticalAnchor = msoAnchorTop
        .MarginLeft = 0
        .MarginRight = Round(size * 0.45, 1)
        .MarginTop = pad
        .MarginBottom = pad
        .TextRange.text = ">>>" & vbCr & "" & vbCr & ">>>" & vbCr & ""
        With .TextRange
            .Font.Name = THEME_FONT
            .Font.size = size
            .Font.Color.RGB = RGB(106, 153, 85)
            With .ParagraphFormat
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
    ' Set AFTER the text. A textbox left to its own devices resizes around what
    ' it contains, and a right-aligned one grows LEFTWARD - which put the
    ' prompts outside the block entirely. SyncGutter pins the same three.
    g.Left = shp.Left
    g.Top = shp.Top
    g.Width = pad + w
    g.Height = shp.Height
    g.ZOrder msoBringToFront
End Sub

' Sara's idea: an ordinary NOTE carrying the output, given the code font and a
' terminal colour so it reads as output rather than as an aside. Costs no new
' machinery at all - the connector already points at the exact line.
Private Function OutputNote(ByVal sld As Slide, ByVal shp As Shape, _
                            ByVal ln As Long, ByVal text As String) As Shape
    Dim n As Shape
    Set n = modNote.AddNote(shp, ln)
    ' Wrap OFF and autofit both ways, so the note hugs the output instead of
    ' folding it. Output is a line the interpreter printed; breaking it across
    ' two lines makes it look like two results.
    With n.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeShapeToFitText
        With .TextRange
            .text = text
            .Font.Name = THEME_FONT
            .Font.size = modBlock.BlockFontSize(shp)
            .Font.Color.RGB = RGB(204, 204, 204)
        End With
    End With
    n.fill.Solid
    n.fill.ForeColor.RGB = RGB(24, 24, 24)
    n.Line.Visible = msoTrue
    n.Line.ForeColor.RGB = RGB(90, 90, 90)
    n.Line.Weight = 1
    Set OutputNote = n
End Function

' Renders every arrow colour at once, top to bottom in list order, so the
' palette can be judged as a set rather than one at a time.
Public Function ArrowPalette(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, i As Long, arr As Shape, lbl As Shape

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    ' Right of the default position, so each arrow gets its full width - which
    ' is the point of the render.
    shp.Left = 180
    shp.Select
    modRibbon.DoStylize

    For i = 0 To ThemeArrowPresetCount() - 1
        modArrow.AddArrow shp, i + 1
    Next i
    StyleBlock shp
    modBlock.UngroupParts shp

    For i = 0 To ThemeArrowPresetCount() - 1
        Set arr = modArrow.FindArrow(shp, i + 1)
        If Not arr Is Nothing Then
            arr.fill.Solid
            arr.fill.ForeColor.RGB = ThemeArrowPreset(i)
            Set lbl = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                            12, arr.Top - 4, 110, 20)
            With lbl.TextFrame.TextRange
                .text = ThemeArrowPresetName(i)
                .Font.size = 14
                .Font.Color.RGB = RGB(60, 60, 60)
                .ParagraphFormat.Alignment = ppAlignRight
            End With
            lbl.Line.Visible = msoFalse
        End If
    Next i

    sld.Export pngPath, "PNG", 1920, 1080
    ArrowPalette = "presets=" & ThemeArrowPresetCount() & vbLf
    Exit Function
Failed:
    ArrowPalette = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Arrows in the left margin: placement, the toggle, and a walkthrough that
' points instead of fading.
Public Function ArrowTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, a As Long, b As Long, arr As Shape
    Dim topPt As Single, hPt As Single, k As Long, marked As Long
    Dim target As Shape, s3 As Shape

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize

    ' Point at line 3, from the cursor.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 3, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoArrow

    r = "arrows=" & modArrow.ArrowCount(shp) & vbLf
    Set arr = modArrow.FindArrow(shp, 3)
    If arr Is Nothing Then
        ArrowTest = r & "ERROR: no arrow"
        Exit Function
    End If

    r = r & "left_of_block=" & Abs(CLng(arr.Left + arr.Width <= shp.Left)) & vbLf
    ' Level with the line's INK, not with the predicted line box.
    If modRender.LineBounds(shp, 3, topPt, hPt) Then
        r = r & "level_with_line=" & _
                Abs(CLng(Abs((arr.Top + arr.Height / 2) - (topPt + hPt / 2)) < 2)) & vbLf
    End If
    ' The code must NOT be faded - that is the whole point of an arrow.
    r = r & "no_emphasis=" & Abs(CLng(Len(modBlock.GetEmphasis(shp)) = 0)) & vbLf

    sld.Export pngPath, "PNG", 1920, 1080

    ' It follows the line when the block is resized.
    Dim wasTop As Single
    wasTop = arr.Top
    shp.Select
    modRibbon.DoSizeUp
    Set arr = modArrow.FindArrow(shp, 3)
    modRender.LineBounds shp, 3, topPt, hPt
    r = r & "followed_resize=" & _
            Abs(CLng(Abs((arr.Top + arr.Height / 2) - (topPt + hPt / 2)) < 2 And _
                     Abs(arr.Top - wasTop) > 1)) & vbLf

    ' Pressing it again on the same line takes it away.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 3, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoArrow
    r = r & "toggled_off=" & Abs(CLng(modArrow.ArrowCount(shp) = 0)) & vbLf

    ' A walkthrough that points instead of fading.
    modOptions.SetStepArrow True
    shp.Select
    modRibbon.DoStepThrough False
    r = r & "slides=" & pres.Slides.count & vbLf
    For k = 2 To pres.Slides.count
        Set target = Nothing
        For Each s3 In modBlock.AllShapes(pres.Slides(k))
            If s3.Tags(modBlock.TAG_BLOCK) = "1" Then Set target = s3
        Next s3
        If Not target Is Nothing Then
            If modArrow.ArrowCount(target) > 0 And Len(modBlock.GetEmphasis(target)) = 0 Then
                marked = marked + 1
            End If
        End If
    Next k
    r = r & "arrowed_slides_unfaded=" & marked & vbLf
    pres.Slides(3).Export Replace(pngPath, ".png", "-walk.png"), "PNG", 1920, 1080

    ' --- the colours ---------------------------------------------------------
    ' The whole claim is that these are the syntax hues darkened until they read
    ' on a WHITE slide, so it is measured. 4.5:1 is stricter than WCAG asks of a
    ' solid shape, deliberately: a washed-out projector eats the margin.
    Dim worst As Double, worstName As String, cc As Double, q As Long
    worst = 999
    For q = 0 To ThemeArrowPresetCount() - 1
        cc = ThemeContrast(ThemeArrowPreset(q), RGB(255, 255, 255))
        If cc < worst Then
            worst = cc
            worstName = ThemeArrowPresetName(q)
        End If
    Next q
    r = r & "worst_on_white=" & Format$(worst, "0.00") & " (" & worstName & ")" & vbLf
    r = r & "all_arrow_colors_read=" & Abs(CLng(worst >= 4.5)) & vbLf

    ' And what the raw palette would have managed, so the darkening is shown to
    ' be doing something rather than assumed to.
    r = r & "raw_class_teal_on_white=" & _
            Format$(ThemeContrast(ThemeColor(tkClass), RGB(255, 255, 255)), "0.00") & vbLf
    r = r & "raw_keyword_purple_on_white=" & _
            Format$(ThemeContrast(ThemeColor(tkKeywordCtrl), RGB(255, 255, 255)), "0.00") & vbLf

    ' Changing the colour repaints every arrow in the DECK, not just this slide.
    modRibbon.DoArrowColor 3                                  ' Green
    Dim wrong As Long, total As Long
    For k = 1 To pres.Slides.count
        For Each s3 In modBlock.AllShapes(pres.Slides(k))
            If Len(s3.Tags(modArrow.TAG_ARROW_OF)) > 0 Then
                total = total + 1
                If s3.fill.ForeColor.RGB <> ThemeArrowPreset(3) Then wrong = wrong + 1
            End If
        Next s3
    Next k
    r = r & "recolored_deck_wide=" & Abs(CLng(total > 1 And wrong = 0)) & _
            " (" & total & " arrows)" & vbLf

    pres.Slides(3).Export Replace(pngPath, ".png", "-colors.png"), "PNG", 1920, 1080

    ' Strip removes them, since an arrow points at styling.
    modOptions.SetStepArrow False
    shp.Select
    modRibbon.DoArrow
    shp.Select
    modRibbon.DoStrip
    r = r & "gone_after_strip=" & Abs(CLng(modArrow.ArrowCount(shp) = 0)) & vbLf

    ArrowTest = r
    Exit Function
Failed:
    ArrowTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' A focused probe: does a connector attached to a note follow the note when the
' note is moved from code? Everything else in NoteTest was too tangled to give a
' clean answer.
Public Function ConnectorProbe(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, a As Long, b As Long, s3 As Shape
    Dim note As Shape, lead As Shape

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, b
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNote

    modBlock.UngroupParts shp
    Set note = modNote.FindNote(shp, 2)
    For Each s3 In modBlock.AllShapes(sld)
        If s3.Tags(modNote.TAG_LEADER_OF) = shp.Tags(modBlock.TAG_ID) Then Set lead = s3
    Next s3

    r = "connected=" & Abs(CLng(lead.ConnectorFormat.BeginConnected = msoTrue And _
                                lead.ConnectorFormat.EndConnected = msoTrue)) & vbLf
    r = r & "note0=" & Format$(note.Left, "0") & "," & Format$(note.Top, "0") & vbLf
    r = r & "lead0=" & Format$(lead.Left, "0") & "," & Format$(lead.Top, "0") & _
            " " & Format$(lead.Width, "0") & "x" & Format$(lead.Height, "0") & vbLf

    note.Left = note.Left + 60
    note.Top = note.Top + 90

    r = r & "note1=" & Format$(note.Left, "0") & "," & Format$(note.Top, "0") & vbLf
    r = r & "lead1=" & Format$(lead.Left, "0") & "," & Format$(lead.Top, "0") & _
            " " & Format$(lead.Width, "0") & "x" & Format$(lead.Height, "0") & vbLf

    lead.RerouteConnections
    r = r & "lead2=" & Format$(lead.Left, "0") & "," & Format$(lead.Top, "0") & _
            " " & Format$(lead.Width, "0") & "x" & Format$(lead.Height, "0") & vbLf

    ' And the same again GROUPED, which is the state a note actually lives in.
    note.Left = note.Left - 60
    note.Top = note.Top - 90
    modBlock.GroupParts shp
    Set note = modNote.FindNote(shp, 2)
    Set lead = Nothing
    For Each s3 In modBlock.AllShapes(sld)
        If s3.Tags(modNote.TAG_LEADER_OF) = shp.Tags(modBlock.TAG_ID) Then Set lead = s3
    Next s3
    r = r & "grouped_note0=" & Format$(note.Left, "0") & "," & Format$(note.Top, "0") & vbLf
    r = r & "grouped_lead0=" & Format$(lead.Width, "0") & "x" & Format$(lead.Height, "0") & vbLf
    note.Left = note.Left + 60
    note.Top = note.Top + 90
    r = r & "grouped_note1=" & Format$(note.Left, "0") & "," & Format$(note.Top, "0") & vbLf
    r = r & "grouped_lead1=" & Format$(lead.Width, "0") & "x" & Format$(lead.Height, "0") & vbLf

    sld.Export pngPath, "PNG", 1920, 1080
    ConnectorProbe = r
    Exit Function
Failed:
    ConnectorProbe = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' The deck-level options: bold emphasis, a note per walkthrough step, the note
' font, and the colour swatches the ribbon shows.
Public Function OptionsTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, a As Long, l As Long, s2 As Shape
    Dim tr As TextRange, colourBefore As Long, k As Long, withNotes As Long
    Dim target As Shape, pic As Object

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize
    Set tr = shp.TextFrame.TextRange

    ' --- bold the emphasised line ------------------------------------------
    ' Build up style: a range, where only the LAST line should go bold.
    modBlock.LineCharRange tr.text, 3, a, l
    colourBefore = tr.Characters(a, 1).Font.Color.RGB

    modBlock.SetEmphasis shp, "2,3"
    modOptions.SetEmphasisBold True
    shp.Select
    modRibbon.DoStylize

    Set tr = shp.TextFrame.TextRange
    modBlock.LineCharRange tr.text, 3, a, l
    r = "last_line_bold=" & Abs(CLng(tr.Characters(a, 1).Font.Bold = msoTrue)) & vbLf
    ' Bold must not collapse the runs the way Font.Highlight did - that bug cost
    ' a whole round of syntax colours once already.
    r = r & "colour_survived_bold=" & _
            Abs(CLng(tr.Characters(a, 1).Font.Color.RGB = colourBefore)) & vbLf
    modBlock.LineCharRange tr.text, 2, a, l
    r = r & "earlier_line_not_bold=" & _
            Abs(CLng(tr.Characters(a, 1).Font.Bold <> msoTrue)) & vbLf

    ' Turning it off has to actually unbold, not leave the last run behind.
    modOptions.SetEmphasisBold False
    shp.Select
    modRibbon.DoStylize
    Set tr = shp.TextFrame.TextRange
    modBlock.LineCharRange tr.text, 3, a, l
    r = r & "bold_cleared=" & Abs(CLng(tr.Characters(a, 1).Font.Bold <> msoTrue)) & vbLf

    ' --- note font ----------------------------------------------------------
    modOptions.SetNoteFont ThemeNoteFontValue(4)          ' Consolas
    modBlock.SetEmphasis shp, "2"
    shp.Select
    modRibbon.DoStylize
    Dim grp As Shape
    Set grp = modBlock.ParentGroup(shp)
    If grp Is Nothing Then shp.Select Else grp.Select
    modRibbon.DoNote
    Set s2 = modNote.FindNote(shp, 2)
    If Not s2 Is Nothing Then
        r = r & "note_font=" & s2.TextFrame.TextRange.Font.Name & vbLf
    End If

    ' --- the shipped default -----------------------------------------------
    ' A fresh deck has no tag, so this is what a note gets before anyone
    ' chooses anything.
    Dim freshPres As Presentation
    Set freshPres = pres
    freshPres.Tags.Delete modOptions.TAG_NOTE_SIZE
    r = r & "default_note_size=" & modOptions.NoteSize() & vbLf
    r = r & "auto_is_distinct=" & _
            Abs(CLng(modOptions.NoteSize() <> modOptions.NOTE_SIZE_AUTO)) & vbLf
    modOptions.SetNoteSize modOptions.NOTE_SIZE_AUTO
    r = r & "auto_survives=" & _
            Abs(CLng(modOptions.NoteSize() = modOptions.NOTE_SIZE_AUTO)) & vbLf
    freshPres.Tags.Delete modOptions.TAG_NOTE_SIZE

    ' --- a swatch for the ribbon -------------------------------------------
    Set pic = modSwatch.Swatch(ThemeNotePreset(1))
    r = r & "swatch_made=" & Abs(CLng(Not pic Is Nothing)) & vbLf
    ' One per preset, since the gallery asks for all of them.
    Dim made As Long, p As Long
    For p = 0 To ThemeNotePresetCount() - 1
        If Not modSwatch.Swatch(ThemeNotePreset(p)) Is Nothing Then made = made + 1
    Next p
    r = r & "swatches=" & made & "/" & ThemeNotePresetCount() & vbLf

    ' --- every preset must be readable --------------------------------------
    ' The claim that a fixed list beats a colour picker rests entirely on this,
    ' so it is measured rather than asserted. 4.5:1 is WCAG AA for body text;
    ' notes are never below 16pt, so this is the strict reading.
    Dim worst As Double, worstName As String, c As Double
    worst = 999
    For p = 0 To ThemeNotePresetCount() - 1
        c = ThemeContrast(ThemeNotePreset(p), ThemeTextOn(ThemeNotePreset(p)))
        If c < worst Then
            worst = c
            worstName = ThemeNotePresetName(p)
        End If
    Next p
    r = r & "worst_contrast=" & Format$(worst, "0.00") & " (" & worstName & ")" & vbLf
    r = r & "all_presets_readable=" & Abs(CLng(worst >= 4.5)) & vbLf

    ' --- two notes, two colours --------------------------------------------
    ' The thing that was broken: restyling reapplied every property from the
    ' deck, so no note could differ from any other in anything.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 3, a, l
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNote

    Dim nA As Shape, nB As Shape

    ' Singled out by the CURSOR, which is the gesture that works on a note
    ' inside a group. Shape.Select on a group member selects the group, so a
    ' test that clicked the note would be testing something a user cannot do.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, l
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNoteColor 9                               ' Emerald
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 3, a, l
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNoteColor 10                              ' Royal

    Set nA = modNote.FindNote(shp, 2)
    Set nB = modNote.FindNote(shp, 3)
    r = r & "two_colours=" & _
            Abs(CLng(nA.fill.ForeColor.RGB = ThemeNotePreset(9) And _
                     nB.fill.ForeColor.RGB = ThemeNotePreset(10))) & vbLf

    ' A per-note size must not drag the other note's size with it.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 2, a, l
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNoteSize 28
    Set nA = modNote.FindNote(shp, 2)
    Set nB = modNote.FindNote(shp, 3)
    r = r & "size_only_selected=" & _
            Abs(CLng(nA.TextFrame.TextRange.Font.size = 28 And _
                     nB.TextFrame.TextRange.Font.size <> 28)) & vbLf
    ' And changing the size must not have reset the colours.
    r = r & "colours_survived_size=" & _
            Abs(CLng(nA.fill.ForeColor.RGB = ThemeNotePreset(9) And _
                     nB.fill.ForeColor.RGB = ThemeNotePreset(10))) & vbLf

    ' With the BLOCK selected and nothing singled out, NOTHING is repainted.
    ' The selection sits on the block after every other command, so a colour
    ' chosen for the next note used to silently repaint the ones already there.
    shp.Select
    modRibbon.DoNoteColor 11                              ' Violet
    Set nA = modNote.FindNote(shp, 2)
    Set nB = modNote.FindNote(shp, 3)
    r = r & "block_selected_repaints_nothing=" & _
            Abs(CLng(nA.fill.ForeColor.RGB = ThemeNotePreset(9) And _
                     nB.fill.ForeColor.RGB = ThemeNotePreset(10))) & vbLf
    ' It did record the choice, which is the whole of what it should have done.
    r = r & "choice_recorded=" & _
            Abs(CLng(modOptions.NoteColor() = ThemeNotePreset(11))) & vbLf

    ' The swatch button applies the recorded colour, but still only to a note
    ' that has been singled out.
    modBlock.LineCharRange shp.TextFrame.TextRange.text, 3, a, l
    shp.TextFrame.TextRange.Characters(a, 1).Select
    modRibbon.DoNoteColorApply
    Set nA = modNote.FindNote(shp, 2)
    Set nB = modNote.FindNote(shp, 3)
    r = r & "swatch_applies_to_cursor_line=" & _
            Abs(CLng(nB.fill.ForeColor.RGB = ThemeNotePreset(11) And _
                     nA.fill.ForeColor.RGB = ThemeNotePreset(9))) & vbLf

    ' Stylize never repaints a note.
    shp.Select
    modRibbon.DoStylize
    Set nA = modNote.FindNote(shp, 2)
    Set nB = modNote.FindNote(shp, 3)
    r = r & "colours_survived_stylize=" & _
            Abs(CLng(nA.fill.ForeColor.RGB = ThemeNotePreset(9) And _
                     nB.fill.ForeColor.RGB = ThemeNotePreset(11))) & vbLf

    sld.Export Replace(pngPath, ".png", "-notes.png"), "PNG", 1920, 1080

    ' --- a note per walkthrough step ---------------------------------------
    ' Bold back on, since the render below is what the two options look like
    ' together, which is how they will actually be used.
    modOptions.SetEmphasisBold True
    modOptions.SetStepNote True
    ' A vibrant preset for the render, since the quiet ones are what the
    ' vibrant half was added to relieve. Through the COMMAND, not the setting:
    ' setting the colour alone leaves notes that already exist as they were,
    ' and the walkthrough would then inherit the old ones.
    shp.Select
    modRibbon.DoNoteColor 6                             ' Crimson
    modBlock.SetEmphasis shp, ""
    shp.Select
    modRibbon.DoStylize
    shp.Select
    modRibbon.DoStepThrough False

    r = r & "slides=" & pres.Slides.count & vbLf
    For k = 2 To pres.Slides.count
        Set target = Nothing
        For Each s2 In modBlock.AllShapes(pres.Slides(k))
            If s2.Tags(modBlock.TAG_BLOCK) = "1" Then Set target = s2
        Next s2
        If Not target Is Nothing Then
            If modNote.NoteCount(target) > 0 Then withNotes = withNotes + 1
        End If
    Next k
    r = r & "step_slides_with_notes=" & withNotes & vbLf

    pres.Slides(3).Export pngPath, "PNG", 1920, 1080

    OptionsTest = r
    Exit Function
Failed:
    OptionsTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Compares three ways of showing an emphasised range, since the text highlight
' stops at the end of each line and looks ragged when several lines are banded.
'   A  text highlight only, as now
'   B  a translucent rectangle over the full width
'   C  no band at all, relying on the dimming
Public Function BandProbe(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape, r As String
    Dim code As String, i As Long, y0 As Single, y1 As Single, rect As Shape
    Dim lineH As Single, pad As Single, size As Single

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H

    For i = 0 To 2
        Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
        sld.Select
        Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
        modBlock.SetEmphasis shp, "1,2,3,5"
        ' C asks the renderer for dimming without a band, rather than removing
        ' the band afterwards - setting Font.Highlight after the colours
        ' collapses the runs and repaints everything one colour.
        If i = 2 Then shp.Tags.Add modBlock.TAG_NOBAND, "1"
        shp.Select
        modRibbon.DoStylize

        If i = 1 Then
            ' B: the band as now, plus a translucent full-width rectangle over
            ' the range. The ragged text band ends up hidden under the rect.
            size = modBlock.BlockFontSize(shp)
            lineH = modSpec.SpecLine(size)
            pad = modSpec.SpecPad(size)
            modBlock.UngroupParts shp
            y0 = shp.Top + pad
            y1 = y0 + 3 * lineH
            Set rect = sld.Shapes.AddShape(msoShapeRectangle, shp.Left + 4, y0, shp.Width - 8, y1 - y0)
            rect.Fill.ForeColor.RGB = RGB(120, 150, 200)
            rect.Fill.Transparency = 0.82
            rect.Line.Visible = msoFalse
            rect.ZOrder msoBringToFront
        End If
    Next i

    pres.Slides(pres.Slides.count - 2).Export Replace(pngPath, ".png", "-A.png"), "PNG", 1920, 1080
    pres.Slides(pres.Slides.count - 1).Export Replace(pngPath, ".png", "-B.png"), "PNG", 1920, 1080
    pres.Slides(pres.Slides.count).Export Replace(pngPath, ".png", "-C.png"), "PNG", 1920, 1080
    BandProbe = "rendered A (text highlight), B (translucent rect), C (dim only)"
    Exit Function
Failed:
    BandProbe = "ERROR " & Err.Number & ": " & Err.Description
End Function

' Builds a walkthrough and checks the slides that come out: how many, in what
' order, and which line each one emphasises.
Public Function StepTest(ByVal srcPath As String, ByVal mode As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim r As String, code As String, i As Long, baseIdx As Long, target As Shape

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize
    baseIdx = sld.SlideIndex
    r = "lines=" & modBlock.CountLines(shp.TextFrame.TextRange.text) & _
        " slides_before=" & pres.Slides.count & vbLf

    shp.Select
    modRibbon.DoStepThrough (mode = "build")
    r = r & "slides_after=" & pres.Slides.count & vbLf
    r = r & "start_slide_emphasis=" & Quoted(modBlock.GetEmphasis(shp)) & vbLf

    For i = baseIdx To pres.Slides.count
        Set target = Nothing
        Dim s2 As Shape
        For Each s2 In modBlock.AllShapes(pres.Slides(i))
            If s2.Tags(modBlock.TAG_BLOCK) = "1" Then Set target = s2
        Next s2
        If Not target Is Nothing Then
            r = r & "  slide " & i & " emphasis=" & Quoted(modBlock.GetEmphasis(target)) & vbLf
        End If
    Next i

    StepTest = r
    Exit Function
Failed:
    StepTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Emphasis, copy and strip, driven through the real commands.
Public Function ToolsTest(ByVal srcPath As String, ByVal pngPath As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim r As String, code As String, txt As String, a As Long, b As Long

    On Error GoTo Failed
    modRibbon.SetQuiet True
    code = ReadTextFile(srcPath)

    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    Set shp = modBlock.CreateBlock(sld, code, modSpec.BASE_SIZE, "python")
    shp.Select
    modRibbon.DoStylize

    ' Emphasise lines 5 and 6 by selecting text inside them.
    txt = shp.TextFrame.TextRange.text
    modBlock.LineCharRange txt, 5, a, b
    Dim endStart As Long, endLen As Long
    modBlock.LineCharRange txt, 6, endStart, endLen
    shp.TextFrame.TextRange.Characters(a, (endStart + endLen) - a).Select
    modRibbon.DoEmphasize
    r = r & "emphasis_tag=" & Quoted(modBlock.GetEmphasis(shp)) & vbLf

    ' It must survive a re-render, which is the entire point.
    shp.Select
    modRibbon.DoStylize
    r = r & "after_rehighlight=" & Quoted(modBlock.GetEmphasis(shp)) & vbLf
    r = r & "line5_banded=" & Banded(shp, 5) & vbLf
    r = r & "line1_banded=" & Banded(shp, 1) & vbLf

    sld.Export pngPath, "PNG", 1920, 1080

    ' Copy, then strip.
    shp.Select
    modRibbon.DoCopyCode
    r = r & "copy_ok=1" & vbLf

    shp.Select
    modRibbon.DoStrip
    r = r & "after_strip_gutter=" & Abs(CLng(modGutter.HasGutter(shp))) & _
            " emphasis=" & Quoted(modBlock.GetEmphasis(shp)) & _
            " guides=" & Abs(CLng(modGuides.GuidesEnabled(shp))) & vbLf

    ToolsTest = r
    Exit Function
Failed:
    ToolsTest = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Whether a line falls inside an emphasis band. The bands are shapes now, so
' this asks geometry: does a band cover this line's row.
Private Function Banded(ByVal shp As Shape, ByVal lineNo As Long) As Long
    Dim sld As Slide, s2 As Shape, y As Single, size As Single
    Banded = 0
    On Error Resume Next
    size = modBlock.BlockFontSize(shp)
    ' The middle of the line's row.
    y = shp.Top + modSpec.SpecPad(size) + (lineNo - 0.5) * modSpec.SpecLine(size)
    Set sld = modGutter.OwningSlide(shp)
    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(modBlock.TAG_BAND_OF) = shp.Tags(modBlock.TAG_ID) Then
            If y >= s2.Top And y <= s2.Top + s2.Height Then Banded = 1
        End If
    Next s2
End Function

' What is actually available for the clipboard and for text backgrounds - both
' are version-dependent, and guessing wrong means a feature that fails on the
' one machine that matters.
Public Function CapabilityProbe(ByVal dummy As String) As String
    Dim r As String, o As Object, pres As Presentation, sld As Slide, shp As Shape

    Set pres = Application.ActivePresentation
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, 40, 40, 400, 100)
    shp.TextFrame.TextRange.text = "abc" & vbCr & "def"

    On Error Resume Next
    Set o = CreateObject("MSForms.DataObject")
    r = r & "MSForms.DataObject: " & IIf(Err.Number = 0, "AVAILABLE", "no (" & Err.Description & ")") & vbLf
    Err.Clear

    shp.TextFrame2.TextRange.Font.Highlight.RGB = RGB(80, 80, 40)
    r = r & "TextFrame2 Font.Highlight: " & IIf(Err.Number = 0, "AVAILABLE", "no (" & Err.Description & ")") & vbLf
    Err.Clear

    shp.TextFrame.TextRange.Copy
    r = r & "TextRange.Copy: " & IIf(Err.Number = 0, "works", "no (" & Err.Description & ")") & vbLf
    Err.Clear

    ' Can we read the user's text selection inside a shape?
    shp.TextFrame.TextRange.Characters(1, 5).Select
    r = r & "Selection.Type after selecting text = " & Application.ActiveWindow.Selection.Type & _
            " (3 = ppSelectionText)" & vbLf
    r = r & "selected text = [" & Application.ActiveWindow.Selection.TextRange.text & "]" & vbLf
    r = r & "sel start=" & Application.ActiveWindow.Selection.TextRange.Start & _
            " len=" & Application.ActiveWindow.Selection.TextRange.Length & vbLf
    On Error GoTo 0

    CapabilityProbe = r
End Function

' Three New block presses in a row. They must not land on top of each other.
Public Function CascadeProbe(ByVal dummy As String) As String
    Dim pres As Presentation, sld As Slide, r As String, i As Long, shp As Shape
    Dim n As Long

    On Error GoTo Failed
    modRibbon.SetQuiet True
    Set pres = Application.ActivePresentation
    pres.PageSetup.SlideWidth = modSpec.SLIDE_W
    pres.PageSetup.SlideHeight = modSpec.SLIDE_H
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)
    sld.Select

    For i = 1 To 3
        modRibbon.DoNewBlock
        n = 0
        For Each shp In modBlock.AllShapes(sld)
            If shp.Tags(modBlock.TAG_BLOCK) = "1" Then n = n + 1
        Next shp
        r = r & "after press " & i & ": blocks_seen=" & n & "  positions="
        For Each shp In modBlock.AllShapes(sld)
            If shp.Tags(modBlock.TAG_BLOCK) = "1" Then
                r = r & "(" & Format$(shp.Left, "0") & "," & Format$(shp.Top, "0") & ")"
            End If
        Next shp
        r = r & vbLf
    Next i

    CascadeProbe = r
    Exit Function
Failed:
    CascadeProbe = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' The soft-line-break case. PowerPoint writes Shift+Enter, and often pasted
' text, as VERTICAL TAB rather than CR. Every count in the add-in must see the
' same visual lines the lexer does, or the numbers and guides drift.
Public Function SoftBreakProbe(ByVal dummy As String) As String
    Dim r As String, txt As String, vt As String
    vt = Chr$(11)

    ' Same six visual lines, written three ways.
    r = r & "all CR      lines=" & modBlock.CountLines("a" & vbCr & "b" & vbCr & "c") & _
            " numbers=[" & Replace(modGutter.NumberColumn("a" & vbCr & "b" & vbCr & "c"), vbCr, "|") & "]" & vbLf
    r = r & "with VT     lines=" & modBlock.CountLines("a" & vt & "b" & vbCr & "c") & _
            " numbers=[" & Replace(modGutter.NumberColumn("a" & vt & "b" & vbCr & "c"), vbCr, "|") & "]" & vbLf

    ' The reported block, with a soft break before the last line.
    txt = "print(1)" & vbCr & "print(2)" & vbCr & "" & vbCr & _
          "for i in x:" & vbCr & "    if y:" & vbCr & "        z()" & vbCr & _
          "print(3)" & vt & "print(4)"
    r = r & "reported    lines=" & modBlock.CountLines(txt) & _
            " numbers=[" & Replace(modGutter.NumberColumn(txt), vbCr, "|") & "]" & vbLf

    Dim levels() As Long, n As Long, i As Long, lv As String
    n = modGuides.IndentLevels(txt, levels)
    For i = 0 To n - 1
        lv = lv & levels(i)
    Next i
    r = r & "            indent levels=" & lv & " (guides must follow these)" & vbLf
    SoftBreakProbe = r
End Function

' Numbering rules, checked without PowerPoint doing any layout: leading and
' trailing blanks unnumbered, everything between numbered consecutively.
Public Function NumberProbe(ByVal dummy As String) As String
    Dim r As String, cases(0 To 4) As String, i As Long, out As String

    cases(0) = "a" & vbCr & "b" & vbCr & "c"
    cases(1) = "a" & vbCr & "b" & vbCr & "c" & vbCr           ' trailing CR
    cases(2) = "" & vbCr & "a" & vbCr & "b" & vbCr & ""       ' padded both ends
    cases(3) = "a" & vbCr & "" & vbCr & "c"                   ' blank in the middle
    cases(4) = "print(1)" & vbCr & "print(2)" & vbCr & "" & vbCr & _
               "for i in x:" & vbCr & "    if y:" & vbCr & "        z()" & vbCr & "done()"

    For i = 0 To 4
        out = Replace(modGutter.NumberColumn(cases(i)), vbCr, "|")
        r = r & "case" & i & " lines=" & modBlock.CountLines(cases(i)) & _
                " numbers=[" & out & "]" & vbLf
    Next i
    NumberProbe = r
End Function

' Compares PowerPoint's own paragraph count against CountLines, for text with
' and without a trailing newline. The reported symptom - the last number appears
' only when there is a blank line after it - is what an off-by-one in the line
' count looks like from the outside.
Public Function CountProbe(ByVal dummy As String) As String
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim r As String, txt As String, i As Long

    On Error GoTo Failed
    Set pres = Application.ActivePresentation
    Set sld = pres.Slides.Add(pres.Slides.count + 1, ppLayoutBlank)

    For i = 0 To 1
        txt = "a = 1" & vbCr & "b = 2" & vbCr & "c = 3"
        If i = 1 Then txt = txt & vbCr          ' with a trailing paragraph mark
        Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, 40, 40 + i * 200, 400, 150)
        shp.TextFrame.WordWrap = msoFalse
        shp.TextFrame.TextRange.text = txt

        r = r & IIf(i = 0, "no trailing CR: ", "trailing CR:    ")
        r = r & "set_len=" & Len(txt) & _
                " readback_len=" & Len(shp.TextFrame.TextRange.text) & _
                " Paragraphs=" & shp.TextFrame.TextRange.Paragraphs.count & _
                " CountLines=" & modBlock.CountLines(shp.TextFrame.TextRange.text) & _
                " lastchar=" & AscW(Right$(shp.TextFrame.TextRange.text, 1)) & vbLf
    Next i

    CountProbe = r
    Exit Function
Failed:
    CountProbe = r & "ERROR " & Err.Number & ": " & Err.Description
End Function

' Isolates the indent-level computation from the drawing, so a failure in one
' cannot be mistaken for the other.
Public Function GuideProbe(ByVal dummy As String) As String
    Dim levels() As Long, n As Long, i As Long, r As String, txt As String
    On Error GoTo Failed
    txt = "for i in range(10):" & vbCr & "    if i == 17:" & vbCr & _
          "        print(1)" & vbCr & vbCr & "    total = 0"
    n = modGuides.IndentLevels(txt, levels)
    r = "lines=" & n & " levels="
    For i = 0 To n - 1
        r = r & levels(i)
    Next i
    GuideProbe = r
    Exit Function
Failed:
    GuideProbe = "ERROR " & Err.Number & ": " & Err.Description
End Function

' One file, for poking at a single sample from the VBA editor's Immediate
' window. The test harness uses RunCorpus instead - see the note there.
' Returns the mask length, or -1 on failure.
Public Function MaskFileToFile(ByVal inPath As String, ByVal outPath As String) As Long
    Dim src As String, mask As String

    On Error GoTo Failed
    src = ReadTextFile(inPath)
    mask = modLexer.MaskOf(src, RunLang())
    WriteTextFile outPath, mask
    MaskFileToFile = Len(mask)
    Exit Function

Failed:
    WriteTextFile outPath, "ERROR " & Err.Number & ": " & Err.Description
    MaskFileToFile = -1
End Function

'------------------------------------------------------------------------------

' Line endings are normalised to LF, matching Python's universal newlines, so
' character positions agree with what the reference tool saw. Without this every
' mask would be one character per line out of step.
Private Function ReadTextFile(ByVal path As String) As String
    Dim f As Integer, buf As String
    f = FreeFile
    Open path For Binary Access Read As #f
    buf = Space$(LOF(f))
    Get #f, , buf
    Close #f
    buf = Replace(buf, vbCrLf, vbLf)
    buf = Replace(buf, vbCr, vbLf)
    ReadTextFile = buf
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, content;
    Close #f
End Sub
