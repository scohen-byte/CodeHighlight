Attribute VB_Name = "modOutput"
'==============================================================================
' modOutput - transcripts: lines that are typed AT an interpreter, and lines
'             that are printed BY one.
'
' Two line kinds, chosen the way emphasis is chosen - select the lines, press
' the button. A line that is neither is ordinary code.
'
'   INTERPRETER  gets the language's prompt, and its code is coloured normally
'   OUTPUT       gets no prompt and no syntax colour, because a printed value
'                is not code and the absence of colour says so
'
' THE PROMPT IS IN THE TEXT, and that is a deliberate reversal.
'
' It was a separate shape, on the same argument the line numbers are made from:
' the block's text is the source, so nothing that is not source should be in it.
' The argument is sound and the result was not. Two shapes that must agree
' LINE FOR LINE will eventually disagree - they have different frames, only one
' has autofit, and a single soft break or a wrapped line puts them out of step -
' and when these two disagree the prompts land on top of the code. The numbers
' survive the same fragility only because they are a narrow right-aligned column
' with a gap, so drift shows up as a wobble rather than as a collision.
'
' In the text there is nothing to keep in step. The prompt indents the code
' because it IS characters in front of the code, and no ruler level, no second
' frame and no alignment nudge is involved.
'
' What that costs is the pure-source property, and Copy code buys it back: it
' strips the prompts and drops the output, so what you paste is what you would
' run. That is a better home for the guarantee anyway - it is exactly the moment
' anybody wants it.
'
' The prompts themselves come from the language table. ">>> " and "... " are
' Python's, and nothing outside modLang*.bas is allowed to know that.
'==============================================================================
Option Explicit

' Comma lists of line numbers, read the same way the emphasis and hidden lists
' are. Both live on the block.
Public Const TAG_OUTPUT As String = "CODEBLOCK_OUTPUT"
Public Const TAG_INTERP As String = "CODEBLOCK_INTERP"

Public Function GetOutputLines(ByVal shp As Shape) As String
    GetOutputLines = shp.Tags(TAG_OUTPUT)
End Function

Public Sub SetOutputLines(ByVal shp As Shape, ByVal lineList As String)
    shp.Tags.Add TAG_OUTPUT, lineList
End Sub

Public Function GetInterpLines(ByVal shp As Shape) As String
    GetInterpLines = shp.Tags(TAG_INTERP)
End Function

Public Sub SetInterpLines(ByVal shp As Shape, ByVal lineList As String)
    shp.Tags.Add TAG_INTERP, lineList
End Sub

Public Function HasOutput(ByVal shp As Shape) As Boolean
    HasOutput = (Len(GetOutputLines(shp)) > 0)
End Function

Public Function IsTranscript(ByVal shp As Shape) As Boolean
    IsTranscript = (Len(GetOutputLines(shp)) > 0 Or Len(GetInterpLines(shp)) > 0)
End Function

' Membership by string search rather than by parsing, the same way the emphasis
' and hidden lists are read. "1,10" must not match line 0 or line 101, which is
' what the commas either side are for.
Public Function InList(ByVal spec As String, ByVal lineNo As Long) As Boolean
    If Len(spec) = 0 Then Exit Function
    InList = (InStr("," & spec & ",", "," & CStr(lineNo) & ",") > 0)
End Function

Public Function IsOutputLine(ByVal spec As String, ByVal lineNo As Long) As Boolean
    IsOutputLine = InList(spec, lineNo)
End Function

' Adds a run of lines to a marking, or takes it out again.
'
' ACCUMULATES, and this is the whole usability of it. A transcript has output in
' several places and a text selection covers only one run at a time, so a
' command that REPLACED the list could never mark both.
'
' Toggling on the whole run rather than per line: if every line in the selection
' is already marked, the gesture plainly means "no it is not".
Public Function ToggleLines(ByVal spec As String, ByVal firstLine As Long, _
                            ByVal lastLine As Long) As String
    Dim marked() As Boolean, hi As Long, i As Long, allOn As Boolean
    Dim parts() As String, v As Long, out As String

    If lastLine < firstLine Then Exit Function

    hi = lastLine
    If Len(spec) > 0 Then
        parts = Split(spec, ",")
        For i = LBound(parts) To UBound(parts)
            v = CLng(Val(parts(i)))
            If v > hi Then hi = v
        Next i
    End If
    If hi < 1 Then Exit Function
    ReDim marked(1 To hi)

    If Len(spec) > 0 Then
        For i = LBound(parts) To UBound(parts)
            v = CLng(Val(parts(i)))
            If v >= 1 And v <= hi Then marked(v) = True
        Next i
    End If

    allOn = True
    For i = firstLine To lastLine
        If Not marked(i) Then
            allOn = False
            Exit For
        End If
    Next i

    For i = firstLine To lastLine
        marked(i) = Not allOn
    Next i

    For i = 1 To hi
        If marked(i) Then
            If Len(out) > 0 Then out = out & ","
            out = out & CStr(i)
        End If
    Next i
    ToggleLines = out
End Function

' Takes a run out of a list without touching the rest, for keeping the two
' markings mutually exclusive - a line cannot be both typed and printed.
Public Function RemoveRun(ByVal spec As String, ByVal firstLine As Long, _
                          ByVal lastLine As Long) As String
    Dim parts() As String, i As Long, v As Long, out As String

    If Len(spec) = 0 Then Exit Function
    parts = Split(spec, ",")
    For i = LBound(parts) To UBound(parts)
        v = CLng(Val(parts(i)))
        If v < firstLine Or v > lastLine Then
            If Len(out) > 0 Then out = out & ","
            out = out & CStr(v)
        End If
    Next i
    RemoveRun = out
End Function

'------------------------------------------------------------------------------
' The prompts, in the text
'------------------------------------------------------------------------------

' Brings every line's prompt into line with the markings. True when the text
' changed, so the caller can reapply the block formatting - assigning to
' TextRange.text drops every run.
'
' Idempotent, because it runs on every Stylize: a line already carrying the
' right prompt is left alone, one carrying the wrong prompt has it swapped, and
' a line that is no longer an interpreter line has its taken off. Without that,
' five Stylizes would give you ">>> >>> >>> >>> >>> x = 1".
Public Function SyncPrompts(ByVal shp As Shape, ByVal langId As String) As Boolean
    Dim lang As LangDef, lines() As String, i As Long, n As Long
    Dim interp As String, want As String, bare As String
    Dim out As String, changed As Boolean, before As String

    lang = modLangRegistry.GetLang(langId)
    If Len(lang.PromptText) = 0 Then Exit Function

    interp = GetInterpLines(shp)
    before = shp.TextFrame.TextRange.text
    If Len(before) = 0 Then Exit Function
    lines = modBlock.SplitLines(before)

    For i = LBound(lines) To UBound(lines)
        n = i - LBound(lines) + 1
        bare = StripPrompt(lines(i), lang)

        want = ""
        If InList(interp, n) Then want = PromptFor(lines, LBound(lines), i, interp, lang)

        If i > LBound(lines) Then out = out & vbCr
        out = out & want & bare
        If want & bare <> lines(i) Then changed = True
    Next i

    If changed Then shp.TextFrame.TextRange.text = out
    SyncPrompts = changed
End Function

' Which prompt one interpreter line should carry.
'
' The continuation prompt goes on the BODY of a statement, and indentation is
' all there is to tell that from without parsing: a line indented past the last
' statement is that statement's body. The base only moves when a line comes back
' out to it, so nesting works.
'
' Measured on the line WITHOUT its prompt, or a line would stop being a
' continuation the moment it was given one.
Private Function PromptFor(ByRef lines() As String, ByVal lo As Long, _
                           ByVal idx As Long, ByVal interp As String, _
                           ByRef lang As LangDef) As String
    Dim j As Long, n As Long, indent As Long, baseIndent As Long
    Dim started As Boolean, bare As String

    For j = lo To idx
        n = j - lo + 1
        If Not InList(interp, n) Then GoTo NextLine

        bare = StripPrompt(lines(j), lang)
        If Len(Trim$(bare)) = 0 Then
            ' A blank line is not a statement, so it gets no prompt. A bare
            ' ">>>" with nothing after it reads as one that failed to render.
            If j = idx Then Exit Function
            GoTo NextLine
        End If

        indent = LeadingColumns(bare)
        If started And indent > baseIndent Then
            If j = idx Then
                PromptFor = lang.ContinueText
                Exit Function
            End If
        Else
            baseIndent = indent
            started = True
            If j = idx Then
                PromptFor = lang.PromptText
                Exit Function
            End If
        End If
NextLine:
    Next j
End Function

' One prompt off the front of a line, whichever of the two it is - and never
' more than one, so a line that has somehow collected two keeps one. Silently
' eating text is worse than leaving something visibly wrong.
Public Function StripPrompt(ByVal line As String, ByRef lang As LangDef) As String
    If Len(lang.PromptText) > 0 Then
        If Left$(line, Len(lang.PromptText)) = lang.PromptText Then
            StripPrompt = Mid$(line, Len(lang.PromptText) + 1)
            Exit Function
        End If
    End If
    If Len(lang.ContinueText) > 0 Then
        If Left$(line, Len(lang.ContinueText)) = lang.ContinueText Then
            StripPrompt = Mid$(line, Len(lang.ContinueText) + 1)
            Exit Function
        End If
    End If
    StripPrompt = line
End Function

' How many characters of prompt a line is carrying, for colouring them.
Public Function PromptLen(ByVal line As String, ByRef lang As LangDef) As Long
    PromptLen = Len(line) - Len(StripPrompt(line, lang))
End Function

' Columns of leading whitespace, counting a tab as a full tab stop.
Private Function LeadingColumns(ByVal line As String) As Long
    Dim i As Long, ch As String, col As Long

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = " " Then
            col = col + 1
        ElseIf ch = vbTab Then
            col = col + CLng(modSpec.TAB_CHARS) - (col Mod CLng(modSpec.TAB_CHARS))
        Else
            Exit For
        End If
    Next i
    LeadingColumns = col
End Function

'------------------------------------------------------------------------------

' What Copy code puts on the clipboard: no prompts, no output.
'
' This is where the pure-source guarantee lives now. Keeping the prompts out of
' the block's text bought that property at every moment EXCEPT the one where
' anybody wants it. Doing it here buys it at exactly that moment.
Public Function CodeOnly(ByVal text As String, ByVal outputSpec As String, _
                         ByVal langId As String) As String
    Dim lang As LangDef, lines() As String, i As Long, n As Long
    Dim out As String, first As Boolean

    lang = modLangRegistry.GetLang(langId)
    lines = modBlock.SplitLines(text)
    first = True

    For i = LBound(lines) To UBound(lines)
        n = i - LBound(lines) + 1
        If Not InList(outputSpec, n) Then
            If Not first Then out = out & vbCr
            out = out & StripPrompt(lines(i), lang)
            first = False
        End If
    Next i
    CodeOnly = out
End Function
