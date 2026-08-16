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

' Drives the REAL ribbon command path - DoNewBlock, DoStylize, DoStylizeAll
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

    ' --- Stylize all --------------------------------------------------------
    modRibbon.DoStylizeAll
    ' AllShapes, not sld.Shapes: Stylize groups a block with its numbers and
    ' guides, so the block is no longer a top-level shape.
    For Each shp In modBlock.AllShapes(sld)
        If modBlock.IsCodeBlock(shp) Then blocks = blocks + 1
    Next shp
    r = r & "highlight_all_blocks=" & blocks & vbLf

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

' Whether a line carries the emphasis band, by COLOUR - Font.Highlight is a
' ColorFormat with no Visible to ask. Answers -1 for "could not tell" rather
' than taking the run down.
Private Function Banded(ByVal shp As Shape, ByVal lineNo As Long) As Long
    Dim a As Long, b As Long, c As Long
    Banded = -1
    On Error Resume Next
    modBlock.LineCharRange shp.TextFrame.TextRange.text, lineNo, a, b
    If a <= 0 Then Exit Function
    If b <= 0 Then b = 1
    c = shp.TextFrame2.TextRange.Characters(a, b).Font.Highlight.RGB
    If Err.Number = 0 Then Banded = Abs(CLng(c = ThemeEmphasisColor()))
    Err.Clear
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
