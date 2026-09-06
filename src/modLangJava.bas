Attribute VB_Name = "modLangJava"
'==============================================================================
' modLangJava - Java syntax table. Data only, no logic.
'==============================================================================
Option Explicit

Public Function JavaLang() As LangDef
    Dim L As LangDef

    L.id = "java"
    L.DisplayName = "Java"
    L.SourceExtension = "java"
    L.PromptText = ""
    L.ContinueText = ""

    L.LineComments = "//"
    L.BlockCommentOpen = "/*"
    L.BlockCommentClose = "*/"

    L.QuoteChars = "'" & """"
    L.TripleQuotes = True             ' Java text blocks use triple double quotes
    L.StringPrefixes = ""
    L.RawPrefixChars = ""
    L.EscapeChar = "\"

    L.InterpPrefixChars = ""
    L.InterpOpen = ""
    L.InterpClose = ""
    L.InterpDoubling = False
    L.InterpSpecChars = ""

    L.NumberPrefixes = "0x 0b"
    L.NumberSuffixes = "l L f F d D"
    L.DigitSep = "_"

    L.CaseSensitive = True
    L.DecoratorChar = "@"
    L.DecoratorLineStartOnly = False  ' annotations may follow modifiers
    L.ExtraIdentChars = "$"
    L.OpenBrackets = "([ {"
    L.OpenBrackets = Replace(L.OpenBrackets, " ", "")
    L.CloseBrackets = ")] }"
    L.CloseBrackets = Replace(L.CloseBrackets, " ", "")

    L.ControlKeywords = "assert break case catch continue default do else " & _
                        "finally for if instanceof new return switch throw " & _
                        "try while yield"

    L.SoftKeywords = ""

    L.DeclKeywords = "abstract class const enum exports extends final goto " & _
                     "implements import interface module native non-sealed " & _
                     "open opens package permits private protected provides " & _
                     "public record requires sealed static strictfp synchronized throws " & _
                     "to transient transitive uses var volatile when with " & _
                     "true false null"

    L.Builtins = ""
    L.TypeNames = "boolean byte char double float int long short void " & _
                  "String Object Class System Math Number Boolean Byte " & _
                  "Character Double Float Integer Long Short Void " & _
                  "Exception RuntimeException Throwable Error Iterable " & _
                  "Collection List ArrayList LinkedList Set HashSet Map " & _
                  "HashMap Optional Stream Arrays Collections"

    L.SelfWords = "this super"
    L.FuncDefKeywords = ""
    L.TypeDefKeywords = "class interface enum record"

    JavaLang = L
End Function
