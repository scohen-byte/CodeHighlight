Attribute VB_Name = "modLexer"
'==============================================================================
' modLexer - the scanner.
'
' A hand-written character scanner. Not regular expressions: a scanner is
' faster, handles multi-line strings and nesting correctly, and can be stepped
' through in the debugger when it gets something wrong.
'
' NOTHING IN THIS MODULE NAMES A LANGUAGE. Every trigger character, keyword set
' and escape rule comes from the LangDef it is handed. Python is one such table,
' in modLangPython. If you find yourself typing "#" or "def" in here, the rule
' belongs in the table instead.
'
' Input:  the raw string from TextRange.Text, plus a LangDef.
' Output: spans of (Start, Length, Kind), Start being 1-based to match
'         TextRange.Characters. Adjacent spans of the same kind are merged as
'         they are emitted, because each one costs a COM round trip later.
'
' Encoding notes that matter:
'   PowerPoint separates paragraphs with CR (13), not CRLF.
'   A soft line break (Shift+Enter) is a vertical tab (11).
'   Both end a line comment. So does LF (10), for text arriving from elsewhere.
'==============================================================================
Option Explicit

Public Type Span
    Start  As Long          ' 1-based, matches TextRange.Characters
    Length As Long
    Kind   As TokenClass
End Type

Private Const SPAN_CHUNK As Long = 256

' --- state for one run --------------------------------------------------------
Private mText As String
Private mLen  As Long
Private mPos  As Long

Private mSpans() As Span
Private mCount   As Long

Private mLang     As LangDef
Private mCtrl     As Collection
Private mDecl     As Collection
Private mBuiltins As Collection
Private mTypes    As Collection
Private mSelfWords As Collection
Private mFuncDef  As Collection
Private mTypeDef  As Collection
Private mPrefixes As Collection

' The word before the current one, used for the "identifier after def" rule.
' Whitespace does not clear it. Anything else does.
Private mPrevWord    As String
Private mAtLineStart As Boolean

'==============================================================================
' Public API
'==============================================================================

' Returns the span count and fills spans(). Returning a count rather than the
' array avoids the zero-length-UDT-array problem, which VBA handles badly.
Public Function Tokenize(ByVal text As String, ByRef lang As LangDef, _
                         ByRef spans() As Span) As Long
    InitRun text, lang
    Do While mPos <= mLen
        ScanOneToken
    Loop
    If mCount > 0 Then
        ReDim Preserve mSpans(0 To mCount - 1)
        spans = mSpans
    End If
    Tokenize = mCount
End Function

' One mask character per source character, lines separated by vbLf. This is the
' format tools/lexref.py emits, so the two can be diffed directly rather than
' compared by eye. See tools/lexref.py for the alphabet.
Public Function MaskOf(ByVal text As String, ByRef lang As LangDef) As String
    Dim spans() As Span, n As Long, i As Long, j As Long
    Dim buf As String, ch As String

    n = Tokenize(text, lang, spans)
    buf = String$(Len(text), ".")

    For i = 0 To n - 1
        ch = ThemeMaskChar(spans(i).Kind)
        For j = spans(i).Start To spans(i).Start + spans(i).Length - 1
            Mid$(buf, j, 1) = ch
        Next j
    Next i

    ' Line breaks stay line breaks, so the mask splits into the same number of
    ' lines as the source.
    For j = 1 To Len(buf)
        If IsLineBreak(Mid$(text, j, 1)) Then Mid$(buf, j, 1) = vbLf
    Next j

    MaskOf = buf
End Function

'==============================================================================
' Setup
'==============================================================================

Private Sub InitRun(ByVal text As String, ByRef lang As LangDef)
    Dim cs As Boolean

    mText = text
    mLen = Len(text)
    mPos = 1
    mCount = 0
    ReDim mSpans(0 To SPAN_CHUNK)

    mLang = lang
    cs = lang.CaseSensitive

    ' Built once per block, not per token. Membership is then a keyed lookup
    ' instead of a scan, and InStr on a keyword blob would match substrings -
    ' "information" would come back as the keyword "in".
    Set mCtrl = WordSet(lang.ControlKeywords, cs)
    Set mDecl = WordSet(lang.DeclKeywords, cs)
    Set mBuiltins = WordSet(lang.Builtins, cs)
    Set mTypes = WordSet(lang.TypeNames, cs)
    Set mSelfWords = WordSet(lang.SelfWords, cs)
    Set mFuncDef = WordSet(lang.FuncDefKeywords, cs)
    Set mTypeDef = WordSet(lang.TypeDefKeywords, cs)
    Set mPrefixes = WordSet(lang.StringPrefixes, False)   ' always case-insensitive

    mPrevWord = ""
    mAtLineStart = True
End Sub

'==============================================================================
' The main dispatch
'==============================================================================

Private Sub ScanOneToken()
    Dim ch As String, marker As String

    ch = Mid$(mText, mPos, 1)

    If IsLineBreak(ch) Then
        Emit mPos, 1, tkDefault
        mPos = mPos + 1
        mPrevWord = ""              ' def / class bind only within a line
        mAtLineStart = True
        Exit Sub
    End If

    ' Whitespace deliberately does NOT clear mPrevWord or mAtLineStart, so
    ' "def  foo" and an indented "@decorator" both still work.
    If ch = " " Or ch = vbTab Then
        Emit mPos, 1, tkDefault
        mPos = mPos + 1
        Exit Sub
    End If

    If MatchesAnyMarker(mLang.LineComments, marker) Then
        ScanLineComment
        Exit Sub
    End If

    If Len(mLang.BlockCommentOpen) > 0 Then
        If Matches(mLang.BlockCommentOpen) Then
            ScanBlockComment
            Exit Sub
        End If
    End If

    If Len(mLang.DecoratorChar) > 0 And mAtLineStart Then
        If ch = mLang.DecoratorChar Then
            ScanDecorator
            Exit Sub
        End If
    End If

    mAtLineStart = False

    If InStr(mLang.QuoteChars, ch) > 0 Then
        ScanString mPos, ""
        Exit Sub
    End If

    If IsDigit(ch) Then
        ScanNumber
        Exit Sub
    End If
    ' A leading dot is a number only when a digit follows: .5 is a float,
    ' but obj.attr is not.
    If ch = "." Then
        If IsDigit(Mid$(mText, mPos + 1, 1)) Then
            ScanNumber
            Exit Sub
        End If
    End If

    If IsIdentStart(ch) Then
        ScanIdentifier
        Exit Sub
    End If

    Emit mPos, 1, tkDefault
    mPos = mPos + 1
    mPrevWord = ""
End Sub

'==============================================================================
' Comments
'==============================================================================

Private Sub ScanLineComment()
    Dim startPos As Long
    startPos = mPos
    Do While mPos <= mLen
        If IsLineBreak(Mid$(mText, mPos, 1)) Then Exit Do
        mPos = mPos + 1
    Loop
    Emit startPos, mPos - startPos, tkComment
    mPrevWord = ""
End Sub

Private Sub ScanBlockComment()
    Dim startPos As Long
    startPos = mPos
    mPos = mPos + Len(mLang.BlockCommentOpen)
    Do While mPos <= mLen
        If Matches(mLang.BlockCommentClose) Then
            mPos = mPos + Len(mLang.BlockCommentClose)
            Exit Do
        End If
        mPos = mPos + 1
    Loop
    Emit startPos, mPos - startPos, tkComment
    mPrevWord = ""
End Sub

'==============================================================================
' Decorators
'==============================================================================

Private Sub ScanDecorator()
    Dim startPos As Long, ch As String
    startPos = mPos
    mPos = mPos + Len(mLang.DecoratorChar)
    ' The dotted name that follows: @functools.wraps, @name.setter
    Do While mPos <= mLen
        ch = Mid$(mText, mPos, 1)
        If IsIdentChar(ch) Or ch = "." Then
            mPos = mPos + 1
        Else
            Exit Do
        End If
    Loop
    Emit startPos, mPos - startPos, tkFunction
    mAtLineStart = False
    mPrevWord = ""
End Sub

'==============================================================================
' Strings
'==============================================================================

' startPos is where the token begins - at the prefix if there is one, so the
' prefix colours as part of the string. mPos is at the opening quote.
Private Sub ScanString(ByVal startPos As Long, ByVal prefix As String)
    Dim quote As String, closer As String, ch As String
    Dim isTriple As Boolean, isRaw As Boolean, isInterp As Boolean
    Dim segStart As Long

    quote = Mid$(mText, mPos, 1)
    isRaw = HasAnyChar(prefix, mLang.RawPrefixChars)
    isInterp = (Len(mLang.InterpOpen) > 0) And _
               HasAnyChar(prefix, mLang.InterpPrefixChars)

    isTriple = False
    If mLang.TripleQuotes Then
        If Mid$(mText, mPos, 3) = quote & quote & quote Then isTriple = True
    End If
    closer = quote
    If isTriple Then closer = quote & quote & quote

    segStart = startPos
    mPos = mPos + Len(closer)

    Do While mPos <= mLen
        ch = Mid$(mText, mPos, 1)

        ' An escape consumes the next character, so a quote cannot end the
        ' string and a trailing backslash cannot swallow the terminator.
        If (Not isRaw) And Len(mLang.EscapeChar) > 0 And ch = mLang.EscapeChar Then
            mPos = mPos + 2
            GoTo ContinueScan
        End If

        If Matches(closer) Then
            mPos = mPos + Len(closer)
            Exit Do
        End If

        ' A single-quoted string never crosses a line. Bailing out here stops
        ' one unterminated quote from colouring the rest of the file.
        If (Not isTriple) And IsLineBreak(ch) Then Exit Do

        If isInterp Then
            If Matches(mLang.InterpOpen) Then
                If mLang.InterpDoubling And _
                   Matches(mLang.InterpOpen & mLang.InterpOpen) Then
                    mPos = mPos + 2 * Len(mLang.InterpOpen)   ' {{ is a literal
                    GoTo ContinueScan
                End If
                ' The brace itself stays string-coloured, the inside is code.
                mPos = mPos + Len(mLang.InterpOpen)
                Emit segStart, mPos - segStart, tkString
                ScanInterpBody
                segStart = mPos          ' spec chars and the closing brace
                GoTo ContinueScan
            End If
        End If

        mPos = mPos + 1
ContinueScan:
    Loop

    Emit segStart, mPos - segStart, tkString
    mPrevWord = ""
End Sub

' Scans the code inside an interpolation, stopping at the closing delimiter or
' at a conversion / format-spec marker. Everything after that marker is part of
' the string again, which is what f"{value:>6}" should look like.
Private Sub ScanInterpBody()
    Dim ch As String, depth As Long

    Do While mPos <= mLen
        ch = Mid$(mText, mPos, 1)

        If depth = 0 Then
            If Matches(mLang.InterpClose) Then Exit Do
            If InStr(mLang.InterpSpecChars, ch) > 0 Then
                ' "!=" is a comparison, not a conversion marker. f"{a != b}"
                ' is legal and common enough to be worth the special case.
                If Not (ch = "!" And Mid$(mText, mPos + 1, 1) = "=") Then Exit Do
            End If
        End If

        ' Bracket depth, so a colon inside a slice or a dict does not read as
        ' the start of a format spec.
        If ch = "(" Or ch = "[" Or ch = "{" Then
            depth = depth + 1
        ElseIf ch = ")" Or ch = "]" Or ch = "}" Then
            depth = depth - 1
        End If

        ' A line break inside an interpolation means something is malformed.
        ' Stop rather than run away.
        If IsLineBreak(ch) Then Exit Do

        ScanOneToken
    Loop
End Sub

'==============================================================================
' Numbers
'==============================================================================

Private Sub ScanNumber()
    Dim startPos As Long, ch As String, nxt As String, prefix As String
    Dim seenDot As Boolean, seenExp As Boolean

    startPos = mPos

    If MatchesNumberPrefix(prefix) Then
        ' 0x / 0o / 0b, then digits of that base plus any separators. Being
        ' loose here is fine - this colours code, it does not validate it.
        mPos = mPos + Len(prefix)
        Do While mPos <= mLen
            ch = Mid$(mText, mPos, 1)
            If IsIdentChar(ch) Or IsDigitSep(ch) Then
                mPos = mPos + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Do While mPos <= mLen
            ch = Mid$(mText, mPos, 1)
            If IsDigit(ch) Or IsDigitSep(ch) Then
                mPos = mPos + 1
            ElseIf ch = "." And (Not seenDot) And (Not seenExp) Then
                seenDot = True
                mPos = mPos + 1
            ElseIf (ch = "e" Or ch = "E") And (Not seenExp) Then
                nxt = Mid$(mText, mPos + 1, 1)
                If IsDigit(nxt) Then
                    seenExp = True
                    mPos = mPos + 1
                ElseIf (nxt = "+" Or nxt = "-") And IsDigit(Mid$(mText, mPos + 2, 1)) Then
                    seenExp = True
                    mPos = mPos + 2
                Else
                    Exit Do          ' a bare e is the start of an identifier
                End If
            Else
                Exit Do
            End If
        Loop
    End If

    If mPos <= mLen Then
        If HasWord(mLang.NumberSuffixes, Mid$(mText, mPos, 1)) Then mPos = mPos + 1
    End If

    Emit startPos, mPos - startPos, tkNumber
    mPrevWord = ""
End Sub

'==============================================================================
' Identifiers
'==============================================================================

Private Sub ScanIdentifier()
    Dim startPos As Long, word As String

    startPos = mPos
    Do While mPos <= mLen
        If Not IsIdentChar(Mid$(mText, mPos, 1)) Then Exit Do
        mPos = mPos + 1
    Loop
    word = Mid$(mText, startPos, mPos - startPos)

    ' A string prefix is not an identifier. rb"..." is one string token, but
    ' "r + 1" really is the variable r, so the quote has to follow immediately.
    If mPos <= mLen Then
        If InStr(mLang.QuoteChars, Mid$(mText, mPos, 1)) > 0 Then
            If InSet(mPrefixes, word, False) Then
                ScanString startPos, word
                Exit Sub
            End If
        End If
    End If

    Emit startPos, mPos - startPos, ClassifyWord(word)
    mPrevWord = word
End Sub

' Order matters. This mirrors tools/lab.py, which is the spec.
Private Function ClassifyWord(ByVal word As String) As TokenClass
    Dim cs As Boolean
    cs = mLang.CaseSensitive

    If InSet(mCtrl, word, cs) Then ClassifyWord = tkKeywordCtrl: Exit Function
    If InSet(mDecl, word, cs) Then ClassifyWord = tkKeywordDecl: Exit Function
    If InSet(mSelfWords, word, cs) Then ClassifyWord = tkKeywordDecl: Exit Function

    If Len(mPrevWord) > 0 Then
        If InSet(mFuncDef, mPrevWord, cs) Then ClassifyWord = tkFunction: Exit Function
        If InSet(mTypeDef, mPrevWord, cs) Then ClassifyWord = tkClass: Exit Function
    End If

    If InSet(mTypes, word, cs) Then ClassifyWord = tkClass: Exit Function

    ' A name followed by "(" is being called, which outranks being a builtin -
    ' that is why len(x) is yellow and a bare int annotation is blue.
    If NextNonBlankIsCall() Then ClassifyWord = tkFunction: Exit Function
    If InSet(mBuiltins, word, cs) Then ClassifyWord = tkKeywordDecl: Exit Function

    ClassifyWord = tkVariable
End Function

Private Function NextNonBlankIsCall() As Boolean
    Dim i As Long, ch As String
    i = mPos
    Do While i <= mLen
        ch = Mid$(mText, i, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        i = i + 1
    Loop
    NextNonBlankIsCall = (Mid$(mText, i, 1) = "(")
End Function

'==============================================================================
' Span accumulation
'==============================================================================

' Merges with the previous span when it is the same kind and directly adjacent.
' Worth doing here rather than in the renderer: every span is a COM round trip,
' and merging typically halves their number.
Private Sub Emit(ByVal startPos As Long, ByVal length As Long, ByVal kind As TokenClass)
    If length <= 0 Then Exit Sub

    If mCount > 0 Then
        If mSpans(mCount - 1).Kind = kind Then
            If mSpans(mCount - 1).Start + mSpans(mCount - 1).Length = startPos Then
                mSpans(mCount - 1).Length = mSpans(mCount - 1).Length + length
                Exit Sub
            End If
        End If
    End If

    If mCount > UBound(mSpans) Then ReDim Preserve mSpans(0 To mCount + SPAN_CHUNK)
    mSpans(mCount).Start = startPos
    mSpans(mCount).Length = length
    mSpans(mCount).Kind = kind
    mCount = mCount + 1
End Sub

'==============================================================================
' Character helpers
'==============================================================================

Private Function Matches(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Matches = (Mid$(mText, mPos, Len(s)) = s)
End Function

Private Function MatchesAnyMarker(ByVal markers As String, ByRef found As String) As Boolean
    Dim parts() As String, i As Long
    If Len(markers) = 0 Then Exit Function
    parts = Split(markers, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Matches(parts(i)) Then
                found = parts(i)
                MatchesAnyMarker = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function MatchesNumberPrefix(ByRef found As String) As Boolean
    Dim parts() As String, i As Long, p As String
    If Len(mLang.NumberPrefixes) = 0 Then Exit Function
    parts = Split(mLang.NumberPrefixes, " ")
    For i = LBound(parts) To UBound(parts)
        p = parts(i)
        If Len(p) > 0 Then
            If LCase$(Mid$(mText, mPos, Len(p))) = LCase$(p) Then
                found = p
                MatchesNumberPrefix = True
                Exit Function
            End If
        End If
    Next i
End Function

' Does any character of chars appear in s, ignoring case. Used for the r and f
' prefix letters, which may be written either way and in either order.
Private Function HasAnyChar(ByVal s As String, ByVal chars As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(chars) = 0 Then Exit Function
    s = LCase$(s)
    chars = LCase$(chars)
    For i = 1 To Len(chars)
        If InStr(s, Mid$(chars, i, 1)) > 0 Then
            HasAnyChar = True
            Exit Function
        End If
    Next i
End Function

Private Function HasWord(ByVal list As String, ByVal word As String) As Boolean
    Dim parts() As String, i As Long
    If Len(list) = 0 Then Exit Function
    parts = Split(list, " ")
    For i = LBound(parts) To UBound(parts)
        If parts(i) = word Then
            HasWord = True
            Exit Function
        End If
    Next i
End Function

Private Function IsDigitSep(ByVal ch As String) As Boolean
    If Len(mLang.DigitSep) = 0 Then Exit Function
    IsDigitSep = (ch = mLang.DigitSep)
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then Exit Function
    c = AscW(ch)
    IsDigit = (c >= 48 And c <= 57)
End Function

Private Function IsIdentStart(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then Exit Function
    c = AscW(ch)
    IsIdentStart = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c = 95)
End Function

Private Function IsIdentChar(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then Exit Function
    c = AscW(ch)
    IsIdentChar = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or _
                  (c = 95) Or (c >= 48 And c <= 57)
End Function

' CR is PowerPoint's paragraph mark, VT is a soft line break, LF turns up in
' text pasted from anywhere else.
Private Function IsLineBreak(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then Exit Function
    c = AscW(ch)
    IsLineBreak = (c = 13 Or c = 10 Or c = 11)
End Function
