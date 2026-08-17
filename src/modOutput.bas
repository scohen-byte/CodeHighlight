Attribute VB_Name = "modOutput"
'==============================================================================
' modOutput - lines INSIDE a block that are output rather than code.
'
' The transcript form: statement, result, statement, result, all in one shape.
' An output NOTE (modNote) is the other shape of the same idea and they are both
' worth having - a note is right when the code is a program and the output is an
' aside about one line of it, and this is right when the slide IS a session at
' the interpreter and the two interleave.
'
' Marked lines are a comma list on a tag, exactly like the hidden lines, because
' it is the same problem: name some lines and treat them differently at render
' time. Marked lines get no syntax colour, no line number and no indent guide,
' and the CODE lines get a prompt.
'
' THE PROMPT IS NOT IN THE BLOCK'S TEXT. It goes in its own shape, the way the
' line numbers do, for the reason the numbers do: the block's text is the source.
' Putting ">>> " in it would break Copy code and feed the lexer four characters
' that are not code.
'
' AND THE PROMPT IS NOT ">>>" AS FAR AS THIS MODULE IS CONCERNED. That string is
' Python's. It comes from the language table, which is the one rule this project
' has held from the start: nothing outside modLang*.bas knows a language.
'==============================================================================
Option Explicit

' Which lines are output, as a comma list, and the tag marking the prompt shape.
Public Const TAG_OUTPUT    As String = "CODEBLOCK_OUTPUT"
Public Const TAG_PROMPT_OF As String = "CODEBLOCK_PROMPT_OF"

Public Function GetOutputLines(ByVal shp As Shape) As String
    GetOutputLines = shp.Tags(TAG_OUTPUT)
End Function

Public Sub SetOutputLines(ByVal shp As Shape, ByVal lineList As String)
    shp.Tags.Add TAG_OUTPUT, lineList
End Sub

Public Function HasOutput(ByVal shp As Shape) As Boolean
    HasOutput = (Len(GetOutputLines(shp)) > 0)
End Function

' Membership by string search rather than by parsing, the same way the emphasis
' and hidden lists are read. "1,10" must not match line 0 or line 101, which is
' what the commas either side are for.
Public Function IsOutputLine(ByVal spec As String, ByVal lineNo As Long) As Boolean
    If Len(spec) = 0 Then Exit Function
    IsOutputLine = (InStr("," & spec & ",", "," & CStr(lineNo) & ",") > 0)
End Function

'------------------------------------------------------------------------------
' The prompt column
'------------------------------------------------------------------------------

Public Function FindPrompt(ByVal shp As Shape) As Shape
    Dim sld As Slide, s2 As Shape, blockId As String

    blockId = shp.Tags(modBlock.TAG_ID)
    If Len(blockId) = 0 Then Exit Function
    Set sld = modGutter.OwningSlide(shp)

    For Each s2 In modBlock.AllShapes(sld)
        If s2.Tags(TAG_PROMPT_OF) = blockId Then
            Set FindPrompt = s2
            Exit Function
        End If
    Next s2
End Function

Public Sub ClearPrompt(ByVal shp As Shape)
    Dim p As Shape
    Set p = FindPrompt(shp)
    If Not p Is Nothing Then p.Delete
End Sub

' The prompt for one line: the language's prompt on a code line, nothing on an
' output line. That column is what makes a transcript readable - without it the
' two kinds of line are only told apart by colour, which a projector may eat.
Public Function PromptColumnText(ByVal text As String, ByVal spec As String, _
                                 ByVal prompt As String) As String
    Dim lines() As String, i As Long, out As String, n As Long

    lines = modBlock.SplitLines(text)
    For i = LBound(lines) To UBound(lines)
        n = i - LBound(lines) + 1
        If i > LBound(lines) Then out = out & vbCr
        If Not IsOutputLine(spec, n) Then out = out & prompt
    Next i
    PromptColumnText = out
End Function

' Creates, updates or removes the prompt column, and owns the block's LEFT
' MARGIN while it exists.
'
' The margin is computed from scratch rather than added to, because this runs on
' every Stylize and adding would walk the code right a little further each time.
Public Sub SyncPrompt(ByVal shp As Shape, ByVal langId As String)
    Dim sld As Slide, p As Shape, lang As LangDef
    Dim size As Single, pad As Single, gutterW As Single, promptW As Single
    Dim lineCount As Long, spec As String, prompt As String

    On Error GoTo Done

    lang = modLangRegistry.GetLang(langId)
    prompt = lang.PromptText
    spec = GetOutputLines(shp)

    size = modBlock.BlockFontSize(shp)
    pad = modSpec.SpecPad(size)
    lineCount = modBlock.CountLines(shp.TextFrame.TextRange.text)
    gutterW = 0
    If modGutter.HasGutter(shp) Then
        gutterW = modSpec.SpecGutter(size, modGutter.FirstLine(shp) + lineCount - 1)
    End If

    ' No output lines, or a language with no prompt, means no column - and the
    ' margin goes back to whatever the gutter alone asks for.
    If Len(spec) = 0 Or Len(prompt) = 0 Then
        ClearPrompt shp
        shp.TextFrame.MarginLeft = pad + gutterW
        Exit Sub
    End If

    promptW = Len(prompt) * modSpec.SpecCharW(size)
    shp.TextFrame.MarginLeft = pad + gutterW + promptW

    Set sld = modGutter.OwningSlide(shp)
    Set p = FindPrompt(shp)
    If p Is Nothing Then
        Set p = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                      shp.Left, shp.Top, promptW, shp.Height)
        p.Tags.Add TAG_PROMPT_OF, shp.Tags(modBlock.TAG_ID)
        p.Line.Visible = msoFalse
        p.fill.Visible = msoFalse
    End If

    With p.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeNone
        .VerticalAnchor = msoAnchorTop
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = pad
        .MarginBottom = pad
        .TextRange.text = PromptColumnText(shp.TextFrame.TextRange.text, spec, prompt)
        With .TextRange
            .Font.Name = THEME_FONT
            .Font.size = size
            .Font.Color.RGB = ThemeOutputMark()
            With .ParagraphFormat
                ' LEFT, unlike the line numbers. A prompt runs into the code it
                ' introduces; a number is a column that lines up on its units.
                .Alignment = ppAlignLeft
                .LineRuleWithin = msoFalse
                .SpaceWithin = modSpec.SpecLine(size)
                .LineRuleBefore = msoFalse
                .SpaceBefore = 0
                .LineRuleAfter = msoFalse
                .SpaceAfter = 0
            End With
        End With
    End With

    ' Pinned AFTER the text. A textbox left to itself resizes around what it
    ' contains, which in the mockup put the prompts outside the block entirely.
    p.Left = shp.Left + pad + gutterW
    p.Top = shp.Top
    p.Width = promptW
    p.Height = modSpec.SpecHeight(size, lineCount)
    If shp.Height > p.Height Then p.Height = shp.Height
    p.ZOrder msoBringToFront

    AlignFirstLine shp, p
Done:
End Sub

' Nudges the column so its first prompt sits on the same baseline as the first
' line of code. The block has autofit on and this does not, and PowerPoint lays
' the two frames out slightly differently as a result - the same correction the
' gutter needs, and for the same reason.
Private Sub AlignFirstLine(ByVal shp As Shape, ByVal p As Shape)
    Dim codeTop As Single, promptTop As Single, delta As Single

    On Error GoTo Done
    If Len(p.TextFrame.TextRange.text) = 0 Then Exit Sub
    If Len(shp.TextFrame.TextRange.text) = 0 Then Exit Sub

    codeTop = shp.TextFrame.TextRange.Characters(1, 1).BoundTop
    promptTop = p.TextFrame.TextRange.Characters(1, 1).BoundTop
    delta = codeTop - promptTop

    If Abs(delta) > 0.2 And Abs(delta) < modSpec.SpecLine(modBlock.BlockFontSize(shp)) Then
        p.Top = p.Top + delta
    End If
Done:
End Sub

'------------------------------------------------------------------------------

' The block's text with the output lines removed, for Copy code.
'
' What you paste should run. A transcript pasted verbatim is not a program, and
' the whole point of Copy code is that it gives you the source rather than a
' picture of a rounded rectangle.
Public Function CodeOnly(ByVal text As String, ByVal spec As String) As String
    Dim lines() As String, i As Long, n As Long, out As String, first As Boolean

    If Len(spec) = 0 Then
        CodeOnly = text
        Exit Function
    End If

    lines = modBlock.SplitLines(text)
    first = True
    For i = LBound(lines) To UBound(lines)
        n = i - LBound(lines) + 1
        If Not IsOutputLine(spec, n) Then
            If Not first Then out = out & vbCr
            out = out & lines(i)
            first = False
        End If
    Next i
    CodeOnly = out
End Function
