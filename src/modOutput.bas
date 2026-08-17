Attribute VB_Name = "modOutput"
'==============================================================================
' modOutput - transcripts: a block that is a session at an interpreter.
'
' THERE IS NO LIST OF OUTPUT LINES, and that is the whole design.
'
' A transcript block carries one flag. Which lines are input and which are
' output is then read off the TEXT: a line carrying a prompt is something you
' typed, a line without one is something the interpreter printed. Nothing is
' stored separately, so nothing can fall out of step.
'
' It was a list of line numbers twice over, and both times the same thing went
' wrong. Line numbers are positions, and positions are invalidated by editing:
' insert a blank line near the top and every marking below it points one line
' too high. Emphasis and hidden lines have the same flaw and get away with it,
' because a band on the wrong line is obvious and you re-mark it. Here the
' prompts are IN the text, so a stale index does not look wrong, it looks like
' a mess - Sara sent three screenshots of exactly that.
'
' Deriving it from the text costs one thing: the prompts have to be right in the
' first place. Turning a block into a transcript prompts every non-blank line,
' and Output takes the prompt back off the lines you select. After that the text
' is the record, and editing it is how you change the record.
'
' What Stylize still recomputes is which of the two prompts a line gets - the
' statement prompt or the continuation one - since that follows indentation and
' indentation changes as you type.
'
' What all this costs is the pure-source property, and Copy code buys it back:
' it strips the prompts, so what you paste is what you would run. That is a
' better home for the guarantee anyway - it is exactly the moment anybody wants
' it, and it gives Copy code a reason to exist beyond convenience.
'
' The prompts come from the language table. ">>> " and "... " are Python's, and
' nothing outside modLang*.bas is allowed to know that.
'==============================================================================
Option Explicit

' One flag on the block. Everything else is read from the text.
Public Const TAG_TRANSCRIPT As String = "CODEBLOCK_TRANSCRIPT"

' Whether this transcript has been given its prompts yet.
'
' A ONE-SHOT, and it has to be, for a reason worth spelling out. A block made as
' a transcript starts empty: you type the session into it, and those lines have
' no prompts, so something must put them there. But "a transcript with no
' prompts gets prompted" cannot be a standing rule - mark every line as Output
' and the next Stylize would put them all back, which is a loop you cannot get
' out of, and an all-output block is a real thing to want (a traceback).
'
' So it fires once, when a transcript first has text to prompt, and after that
' the text is the record and a newly typed line stays bare until you say
' otherwise.
Public Const TAG_PROMPTED As String = "CODEBLOCK_PROMPTED"

Public Function IsTranscript(ByVal shp As Shape) As Boolean
    IsTranscript = (shp.Tags(TAG_TRANSCRIPT) = "1")
End Function

Public Sub SetTranscript(ByVal shp As Shape, ByVal on_ As Boolean)
    shp.Tags.Add TAG_TRANSCRIPT, IIf(on_, "1", "")
    ' Turning it off forgets that it was ever prompted, so turning it back on
    ' prompts again rather than leaving a bare block looking like all output.
    If Not on_ Then shp.Tags.Add TAG_PROMPTED, ""
End Sub

'------------------------------------------------------------------------------
' The prompts, in the text
'------------------------------------------------------------------------------

' Recomputes WHICH prompt each prompted line carries, and nothing else.
'
' A line that has no prompt is output and is left alone - that is how the text
' records the marking. A line that has one keeps one, but it may swap between
' the statement prompt and the continuation prompt, since that follows the
' indentation and the indentation changes as you type.
'
' True when the text changed, so the caller can reapply the block formatting:
' assigning to TextRange.text drops every run.
Public Function SyncPrompts(ByVal shp As Shape, ByVal langId As String) As Boolean
    Dim lang As LangDef, lines() As String, i As Long
    Dim want As String, bare As String, out As String
    Dim changed As Boolean, before As String, prompted() As Boolean, n As Long

    lang = modLangRegistry.GetLang(langId)
    If Len(lang.PromptText) = 0 Then Exit Function

    before = shp.TextFrame.TextRange.text
    If Len(before) = 0 Then Exit Function
    lines = modBlock.SplitLines(before)
    n = UBound(lines) - LBound(lines) + 1
    If n < 1 Then Exit Function

    ' First text a transcript has ever had: prompt all of it. See TAG_PROMPTED.
    If IsTranscript(shp) And shp.Tags(TAG_PROMPTED) <> "1" Then
        If AnyContent(lines) Then
            shp.Tags.Add TAG_PROMPTED, "1"
            SyncPrompts = SetAllPrompts(shp, langId, True)
            Exit Function
        End If
    End If

    ReDim prompted(1 To n)
    For i = LBound(lines) To UBound(lines)
        prompted(i - LBound(lines) + 1) = _
            (PromptLen(lines(i), lang) > 0) And IsTranscript(shp)
    Next i

    For i = LBound(lines) To UBound(lines)
        bare = StripPrompt(lines(i), lang)
        want = ""
        If prompted(i - LBound(lines) + 1) Then
            want = PromptFor(lines, LBound(lines), i, prompted, lang)
        End If

        If i > LBound(lines) Then out = out & vbCr
        out = out & want & bare
        If want & bare <> lines(i) Then changed = True
    Next i

    If changed Then shp.TextFrame.TextRange.text = out
    SyncPrompts = changed
End Function

Private Function AnyContent(ByRef lines() As String) As Boolean
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            AnyContent = True
            Exit Function
        End If
    Next i
End Function

' Puts a prompt on every non-blank line, or takes them all off. This is the one
' moment the marking is CREATED - after it, the text is the record.
'
' It owns TAG_PROMPTED, and every caller therefore agrees about it. Leaving that
' to the callers meant a block turned into a transcript by pressing Output was
' prompted here, had the one line's prompt taken off, and was then prompted all
' over again by the one-shot rule on the very next Stylize - which undid the
' thing the press was for.
Public Function SetAllPrompts(ByVal shp As Shape, ByVal langId As String, _
                              ByVal on_ As Boolean) As Boolean
    Dim lang As LangDef, lines() As String, i As Long
    Dim bare As String, out As String, changed As Boolean

    lang = modLangRegistry.GetLang(langId)
    If Len(lang.PromptText) = 0 Then Exit Function
    lines = modBlock.SplitLines(shp.TextFrame.TextRange.text)

    For i = LBound(lines) To UBound(lines)
        bare = StripPrompt(lines(i), lang)
        If i > LBound(lines) Then out = out & vbCr
        ' A blank line is not a statement, so it never gets a prompt: a bare
        ' ">>>" with nothing after it reads as one that failed to render.
        If on_ And Len(Trim$(bare)) > 0 Then
            out = out & lang.PromptText & bare
        Else
            out = out & bare
        End If
        If out <> "" Then changed = True
    Next i

    changed = (out <> shp.TextFrame.TextRange.text)
    If changed Then shp.TextFrame.TextRange.text = out
    shp.Tags.Add TAG_PROMPTED, IIf(on_, "1", "")
    SetAllPrompts = changed
End Function

' Puts a prompt on a run of lines, or takes it off. NOT a toggle.
'
' Two explicit commands rather than one that flips. A toggle reads fine on a
' button called Output while it is taking prompts OFF, and reads as nonsense
' the moment it is putting them back - "press Output to make this a line you
' typed" is not a sentence anybody should have to work out. And a new line
' starts bare, which is to say output, so putting a prompt on it is a thing you
' do often enough to deserve its own button.
Public Function SetPrompts(ByVal shp As Shape, ByVal langId As String, _
                           ByVal ln1 As Long, ByVal ln2 As Long, _
                           ByVal on_ As Boolean) As Boolean
    Dim lang As LangDef, lines() As String, i As Long, n As Long
    Dim bare As String, out As String

    lang = modLangRegistry.GetLang(langId)
    If Len(lang.PromptText) = 0 Then Exit Function
    lines = modBlock.SplitLines(shp.TextFrame.TextRange.text)

    For i = LBound(lines) To UBound(lines)
        n = i - LBound(lines) + 1
        If i > LBound(lines) Then out = out & vbCr

        If n < ln1 Or n > ln2 Then
            out = out & lines(i)
        Else
            bare = StripPrompt(lines(i), lang)
            ' A blank line never gets a prompt: a bare ">>>" with nothing after
            ' it reads as a statement that failed to render.
            If on_ And Len(Trim$(bare)) > 0 Then
                out = out & lang.PromptText & bare
            Else
                out = out & bare
            End If
        End If
    Next i

    If out <> shp.TextFrame.TextRange.text Then
        shp.TextFrame.TextRange.text = out
        SetPrompts = True
    End If
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
                           ByVal idx As Long, ByRef prompted() As Boolean, _
                           ByRef lang As LangDef) As String
    Dim j As Long, n As Long, indent As Long, baseIndent As Long
    Dim started As Boolean, bare As String

    For j = lo To idx
        n = j - lo + 1
        If Not prompted(n) Then GoTo NextLine

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
Public Function CodeOnly(ByVal text As String, ByVal langId As String) As String
    Dim lang As LangDef, lines() As String, i As Long
    Dim out As String, first As Boolean

    lang = modLangRegistry.GetLang(langId)
    lines = modBlock.SplitLines(text)
    first = True

    For i = LBound(lines) To UBound(lines)
        ' A prompted line is one you typed, so it survives with its prompt
        ' stripped. A bare one is output, and output is not code.
        If PromptLen(lines(i), lang) > 0 Then
            If Not first Then out = out & vbCr
            out = out & StripPrompt(lines(i), lang)
            first = False
        End If
    Next i
    CodeOnly = out
End Function
