Attribute VB_Name = "modLangRegistry"
'==============================================================================
' modLangRegistry - the language table, and the only place that knows which
'                   languages exist.
'
' The lexer is generic. It scans comments, strings, numbers, identifiers and
' operators using the rules in a LangDef, and never mentions a language by name.
' Everything language-specific lives in a table like modLangPython.PythonLang().
'
' TO ADD A LANGUAGE:
'   1. Copy modLangPython.bas to modLangJava.bas and fill in the table.
'   2. Add one Register line to BuildRegistry below.
' Nothing else changes. Not the lexer, not the renderer, not the ribbon XML,
' not the palette - token classes are language-neutral by design.
'==============================================================================
Option Explicit

' Word lists are space-delimited strings here and expanded into Collections
' once at load. Never test membership with InStr on the raw string: it matches
' substrings, so "information" would come back as the keyword "in".
Public Type LangDef
    Id                As String   ' "python" - stable, stored in the shape tag
    DisplayName       As String   ' "Python" - what the ribbon dropdown shows

    ' --- comments -----------------------------------------------------------
    LineComments      As String   ' space-delimited markers, e.g. "#" or "// #"
    BlockCommentOpen  As String   ' "" when the language has none
    BlockCommentClose As String

    ' --- strings ------------------------------------------------------------
    QuoteChars        As String   ' every char that opens a string, e.g. "'"""
    TripleQuotes      As Boolean  ' does ''' / """ span lines
    StringPrefixes    As String   ' space-delimited, case-insensitive: "r b u f rb br fr rf"
    RawPrefixChars    As String   ' prefix letters that disable escapes, e.g. "r"
    EscapeChar        As String   ' "\" - empty disables escape handling

    ' --- interpolation ------------------------------------------------------
    ' Python f-strings, and the same shape covers JS template literals. Empty
    ' InterpOpen disables the whole thing, so a language without interpolation
    ' costs nothing.
    InterpPrefixChars As String   ' prefix letters that enable it, "f"
    InterpOpen        As String   ' "{"   (JS would be "${")
    InterpClose       As String   ' "}"
    InterpDoubling    As Boolean  ' is "{{" an escaped literal brace
    InterpSpecChars   As String   ' at depth 0 these end the code part: "!:"

    ' --- numbers ------------------------------------------------------------
    NumberPrefixes    As String   ' "0x 0o 0b"
    NumberSuffixes    As String   ' "j J" - trailing type markers
    DigitSep          As String   ' "_" - empty if the language has none

    ' --- the interpreter ----------------------------------------------------
    ' The prompt a REPL prints before a line you type, INCLUDING its trailing
    ' space: ">>> " for Python, "> " for a shell, "" for a language with no
    ' REPL. modOutput draws it beside code lines in a transcript, and empty
    ' means a transcript simply gets no prompt column.
    PromptText        As String
    ' What a REPL prints before a line that CONTINUES the previous one - the
    ' body of a loop typed at the interpreter. "... " for Python. Empty falls
    ' back to blank, which is right for a language with no such thing.
    ContinueText      As String

    ' --- identifiers --------------------------------------------------------
    CaseSensitive     As Boolean
    DecoratorChar     As String   ' "@" at line start, "" if the language has none
    OpenBrackets      As String   ' "([{" - empty disables depth colouring
    CloseBrackets     As String   ' ")]}" - must pair positionally with the above
    ControlKeywords   As String   ' purple - flow control
    SoftKeywords      As String   ' purple, but only where a statement can start
    DeclKeywords      As String   ' blue   - declarations, literals, word operators
    Builtins          As String   ' yellow when called
    TypeNames         As String   ' teal   - known types
    SelfWords         As String   ' blue   - "self cls" / "this"
    FuncDefKeywords   As String   ' identifier right after one of these is a function
    TypeDefKeywords   As String   ' identifier right after one of these is a type
End Type

Private mLangs()    As LangDef
Private mCount      As Long
Private mBuilt      As Boolean

'------------------------------------------------------------------------------
' Registry
'------------------------------------------------------------------------------

Private Sub BuildRegistry()
    If mBuilt Then Exit Sub
    mCount = 0
    ReDim mLangs(0 To 7)

    Register modLangPython.PythonLang()
    ' Register modLangJava.JavaLang()
    ' Register modLangC.CLang()

    mBuilt = True
End Sub

Private Sub Register(ByRef lang As LangDef)
    If mCount > UBound(mLangs) Then ReDim Preserve mLangs(0 To mCount + 7)
    mLangs(mCount) = lang
    mCount = mCount + 1
End Sub

'------------------------------------------------------------------------------
' Lookup
'------------------------------------------------------------------------------

Public Function LangCount() As Long
    BuildRegistry
    LangCount = mCount
End Function

Public Function LangAt(ByVal index As Long) As LangDef
    BuildRegistry
    If index < 0 Or index >= mCount Then index = 0
    LangAt = mLangs(index)
End Function

Public Function LangIndexOf(ByVal id As String) As Long
    Dim i As Long
    BuildRegistry
    For i = 0 To mCount - 1
        If StrComp(mLangs(i).id, id, vbTextCompare) = 0 Then
            LangIndexOf = i
            Exit Function
        End If
    Next i
    LangIndexOf = -1        ' caller decides whether to fall back to the default
End Function

' Unknown ids fall back to the default language rather than raising. A block
' pasted in from a future version of the add-in should still highlight as
' something sensible instead of erroring.
Public Function GetLang(ByVal id As String) As LangDef
    Dim i As Long
    i = LangIndexOf(id)
    If i < 0 Then i = 0
    GetLang = LangAt(i)
End Function

Public Function DefaultLangId() As String
    DefaultLangId = LangAt(0).id
End Function

'------------------------------------------------------------------------------
' Word-list helpers, used by the lexer
'------------------------------------------------------------------------------

' Expands a space-delimited list into a Collection keyed by word, so membership
' is a keyed lookup rather than a scan. Built once per block, not per token.
Public Function WordSet(ByVal words As String, ByVal caseSensitive As Boolean) As Collection
    Dim c As Collection, parts() As String, i As Long, w As String
    Set c = New Collection
    parts = Split(Trim$(words), " ")
    For i = LBound(parts) To UBound(parts)
        w = Trim$(parts(i))
        If Len(w) > 0 Then
            If Not caseSensitive Then w = LCase$(w)
            On Error Resume Next     ' duplicates in a table are harmless
            c.Add w, w
            On Error GoTo 0
        End If
    Next i
    Set WordSet = c
End Function

Public Function InSet(ByVal words As Collection, ByVal word As String, _
                      ByVal caseSensitive As Boolean) As Boolean
    Dim dummy As String
    If Not caseSensitive Then word = LCase$(word)
    On Error Resume Next
    dummy = words.Item(word)
    InSet = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
