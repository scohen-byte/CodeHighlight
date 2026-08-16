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
    L.OpenBrackets = "([{"
    L.CloseBrackets = ")]}"

    ' Purple. Flow control.
    L.ControlKeywords = "if elif else for while return break continue " & _
                        "try except finally raise with as yield pass " & _
                        "assert del import from global nonlocal async " & _
                        "await " & _
                        "and or not in is"

    ' Soft keywords: real keywords only at the start of a statement. Elsewhere
    ' they are ordinary names, and re.match(...) or a variable called match is
    ' far more common in teaching code than a match statement.
    L.SoftKeywords = "match case"

    ' Blue. Declarations and literals.
    '
    ' and / or / not / in / is are NOT here - measured against real VS Code, they
    ' render purple with the control keywords, not blue. pygments files them
    ' under Operator, which is what led the reference astray.
    L.DeclKeywords = "def class lambda True False None"

    ' Yellow when followed by "(", blue otherwise.
    ' Builtin FUNCTIONS only. The builtin classes live in TypeNames above and
    ' are colour-checked before this list is consulted.
    L.Builtins = "print len sum min max abs sorted open input isinstance " & _
                 "issubclass repr round any all next iter format hash id " & _
                 "dir vars getattr setattr hasattr delattr callable divmod " & _
                 "pow bin hex oct chr ord"

    ' Teal. Builtin CLASSES, which VS Code colours as types even when called -
    ' measured: list(x) renders #4EC9B0, not function yellow.
    '
    ' This was empty until 2026-08-16, because the reference implementation uses
    ' pygments, which has no semantic model and calls list() a function. Real VS
    ' Code runs Pylance, which knows it is a class. Where the two disagree, the
    ' editor wins - matching what a student sees is the whole point.
    '
    ' Everything below is a class in CPython, including the ones that read like
    ' functions: enumerate, zip, map, filter and reversed are all types.
    L.TypeNames = "int str float bool complex bytes bytearray memoryview " & _
                  "list tuple dict set frozenset range slice object type " & _
                  "enumerate zip map filter reversed property " & _
                  "staticmethod classmethod super " & _
                  "Exception BaseException ValueError TypeError KeyError " & _
                  "IndexError AttributeError RuntimeError StopIteration " & _
                  "NotImplementedError ZeroDivisionError FileNotFoundError " & _
                  "OSError IOError ImportError NameError SyntaxError " & _
                  "IndentationError KeyboardInterrupt MemoryError " & _
                  "OverflowError RecursionError AssertionError " & _
                  "ArithmeticError LookupError PermissionError " & _
                  "UnicodeDecodeError UnicodeEncodeError"

    L.SelfWords = "self cls"

    L.FuncDefKeywords = "def"
    L.TypeDefKeywords = "class"

    PythonLang = L
End Function
