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

Public Sub DoHighlight()
    Dim shp As Shape, problem As String, langId As String
    Dim before As String, after As String, size As Single

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

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

    modRender.ApplyHighlight shp, langId
    modBlock.ResizeToContent shp
    ' After the resize, so the guides are drawn against final geometry.
    modGuides.DrawGuides shp
End Sub

Public Sub DoHighlightAll()
    Dim sld As Slide, shp As Shape, count As Long

    Set sld = ActiveSlide()
    If sld Is Nothing Then
        Warn "Open a slide in Normal view first."
        Exit Sub
    End If

    For Each shp In sld.Shapes
        If modBlock.IsCodeBlock(shp) And shp.HasTextFrame Then
            modRender.ApplyHighlight shp, modBlock.BlockLangId(shp, CurrentLangId())
            modBlock.ResizeToContent shp
            modGuides.DrawGuides shp
            count = count + 1
        End If
    Next shp

    If count = 0 Then
        Warn "No code blocks on this slide. Highlight one first, which tags it."
    End If
End Sub

Public Sub DoSizeUp()
    StepSize 1
End Sub

Public Sub DoSizeDown()
    StepSize -1
End Sub

' Picks the largest rung the code fits at, and says so when that is too small
' to read from the back of a room rather than silently shrinking it.
Public Sub DoFit()
    Dim shp As Shape, problem As String, best As Single, slides As Long

    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then
        Warn problem
        Exit Sub
    End If

    best = FitSizeFor(shp)
    modBlock.ApplySize shp, best
    modGuides.DrawGuides shp

    If best < modSpec.MIN_TEACHING_SIZE Then
        slides = SlidesNeededFor(shp)
        Warn "This only fits at " & Format$(best, "0") & "pt, which will not read " & _
             "from the back of a lecture hall." & vbCrLf & vbCrLf & _
             "At " & Format$(modSpec.MIN_TEACHING_SIZE, "0") & "pt it needs about " & _
             slides & " slide" & IIf(slides = 1, "", "s") & ". Consider splitting it."
    End If
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

    modBlock.ApplySize shp, newSize
    modGuides.DrawGuides shp
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

' Phase 3 will answer this from the gutter shape. Until then a block never has
' one, and Fit is very slightly generous with the width as a result.
Private Function HasGutter(ByVal shp As Shape) As Boolean
    HasGutter = False
End Function

Public Sub DoSetLanguage(ByVal index As Long)
    Dim shp As Shape, problem As String

    mLangIndex = index

    ' If a block is selected, the choice applies to it and takes effect at once.
    Set shp = modBlock.SelectedBlock(problem)
    If shp Is Nothing Then Exit Sub

    modBlock.SetBlockLang shp, CurrentLangId()
    modRender.ApplyHighlight shp, CurrentLangId()
End Sub

'------------------------------------------------------------------------------
' Ribbon callbacks
'------------------------------------------------------------------------------

Public Sub RibbonNewBlock(control As IRibbonControl)
    DoNewBlock
End Sub

Public Sub RibbonHighlight(control As IRibbonControl)
    DoHighlight
End Sub

Public Sub RibbonHighlightAll(control As IRibbonControl)
    DoHighlightAll
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
    returnedVal = False                ' Phase 3
End Sub

Public Sub RibbonToggleGutter(control As IRibbonControl, pressed As Boolean)
    NotYet "Line numbers"              ' Phase 3
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

Private Sub Warn(ByVal msg As String)
    mLastWarning = msg
    If Not mQuiet Then MsgBox msg, vbExclamation, ADDIN_NAME
End Sub

Private Sub NotYet(ByVal what As String)
    Warn what & " is not built yet."
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
