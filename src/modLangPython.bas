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

    L.NumberPrefixes = "0x 0o 0b"
    L.NumberSuffixes = "j J"
    L.DigitSep = "_"

    L.CaseSensitive = True
    L.DecoratorChar = "@"

    ' Purple. Flow control.
    L.ControlKeywords = "if elif else for while return break continue " & _
                        "try except finally raise with as yield pass " & _
                        "assert del import from global nonlocal async " & _
                        "await match case"

    ' Blue. Declarations, literals, and the word operators - pygments files
    ' and/or/not/in/is under Operator, but VS Code renders them blue like these.
    L.DeclKeywords = "def class lambda True False None and or not in is"

    ' Yellow when followed by "(", blue otherwise.
    L.Builtins = "print len range int str float list dict set tuple bool " & _
                 "enumerate zip map filter sum min max abs sorted reversed " & _
                 "open input isinstance issubclass type super repr round " & _
                 "any all next iter format hash id dir vars getattr setattr " & _
                 "hasattr delattr callable divmod pow bin hex oct chr ord"

    ' Teal.
    L.TypeNames = "object Exception BaseException ValueError TypeError " & _
                  "KeyError IndexError AttributeError RuntimeError " & _
                  "StopIteration NotImplementedError ZeroDivisionError " & _
                  "FileNotFoundError OSError IOError"

    L.SelfWords = "self cls"

    L.FuncDefKeywords = "def"
    L.TypeDefKeywords = "class"

    PythonLang = L
End Function
