Attribute VB_Name = "modLangPython"
'==============================================================================
' modLangPython - the Python table. Data only, no logic.
'
' This is the template for every other language. A new language is a copy of
' this file with the fields filled in differently, plus one Register line in
' modLangRegistry.BuildRegistry.
'
' Word lists mirror tools/lab.py, which is the reference implementation of the
' locked spec. When this file and lab.py disagree, lab.py is right.
'==============================================================================
Option Explicit

Public Function PythonLang() As LangDef
    Dim L As LangDef

    L.id = "python"
    L.DisplayName = "Python"

    L.LineComments = "#"
    L.BlockCommentOpen = ""            ' Python has none - docstrings are strings
    L.BlockCommentClose = ""

    L.QuoteChars = "'" & """"
    L.TripleQuotes = True
    L.StringPrefixes = "r b u f rb br fr rf"
    L.RawPrefixChars = "r"
    L.EscapeChar = "\"

    L.InterpPrefixChars = "f"
    L.InterpOpen = "{"
    L.InterpClose = "}"
    L.InterpDoubling = True        ' "{{" is a literal brace, not a field
    L.InterpSpecChars = "!:"       ' conversion and format spec end the code part

    L.NumberPrefixes = "0x 0o 0b"
    L.NumberSuffixes = "j J"
    L.DigitSep = "_"

    L.CaseSensitive = True
    L.DecoratorChar = "@"

    ' Purple. Flow control.
    L.ControlKeywords = "if elif else for while return break continue " & _
                        "try except finally raise with as yield pass " & _
                        "assert del import from global nonlocal async " & _
                        "await"

    ' Soft keywords: real keywords only at the start of a statement. Elsewhere
    ' they are ordinary names, and re.match(...) or a variable called match is
    ' far more common in teaching code than a match statement.
    L.SoftKeywords = "match case"

    ' Blue. Declarations, literals, and the word operators - pygments files
    ' and/or/not/in/is under Operator, but VS Code renders them blue like these.
    L.DeclKeywords = "def class lambda True False None and or not in is"

    ' Yellow when followed by "(", blue otherwise.
    L.Builtins = "print len range int str float list dict set tuple bool " & _
                 "enumerate zip map filter sum min max abs sorted reversed " & _
                 "open input isinstance issubclass type super repr round " & _
                 "any all next iter format hash id dir vars getattr setattr " & _
                 "hasattr delattr callable divmod pow bin hex oct chr ord"

    ' Teal. DELIBERATELY EMPTY for Python.
    '
    ' Tempting to put the builtin exceptions here, but the signed-off look
    ' reserves teal for the name in a class DEFINITION. `ValueError(...)` is
    ' yellow, because it is being called. Filling this in would diverge from
    ' tools/lab.py, which is the spec.
    '
    ' The field stays because it earns its place in other languages - Java's
    ' String and int genuinely are types worth colouring as types.
    L.TypeNames = ""

    L.SelfWords = "self cls"

    L.FuncDefKeywords = "def"
    L.TypeDefKeywords = "class"

    PythonLang = L
End Function
