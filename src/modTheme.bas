Attribute VB_Name = "modTheme"
'==============================================================================
' modTheme - token classes and their colours.
'
' Token classes are LANGUAGE-NEUTRAL. Comment, String, Number, Keyword,
' Function, Class and Variable are concepts every language in scope shares, so
' adding a language never touches this module. That is the whole reason the
' palette sits behind a function instead of being inlined in the renderer.
'
' Colours are VS Code Dark Modern, matching tools/lab.py exactly. lab.py is the
' spec - when this disagrees with it, this is wrong.
'
' Always build colours with RGB(). VBA stores a colour as BGR internally, so a
' raw &H literal copied from a hex colour comes out with red and blue swapped,
' which looks almost right and is maddening to spot.
'==============================================================================
Option Explicit

Public Enum TokenClass
    tkDefault = 0           ' operators, punctuation
    tkComment = 1
    tkString = 2
    tkNumber = 3
    tkKeywordCtrl = 4       ' flow control
    tkKeywordDecl = 5       ' declarations, literals, word operators, self
    tkFunction = 6          ' definitions, call sites, builtins, decorators
    tkClass = 7             ' the name in a class definition
    tkVariable = 8          ' plain identifiers, parameters, attributes
    ' Bracket pairs, coloured by nesting depth and cycling every three levels,
    ' the way VS Code does by default. Language-neutral like every other class -
    ' which brackets exist is a property of the language table.
    tkBracket1 = 9
    tkBracket2 = 10
    tkBracket3 = 11
End Enum

Public Const THEME_FONT As String = "Consolas"

Public Function ThemeColor(ByVal kind As TokenClass) As Long
    Select Case kind
        Case tkComment:      ThemeColor = RGB(106, 153, 85)     ' 6A9955
        Case tkString:       ThemeColor = RGB(206, 145, 120)    ' CE9178
        Case tkNumber:       ThemeColor = RGB(181, 206, 168)    ' B5CEA8
        Case tkKeywordCtrl:  ThemeColor = RGB(197, 134, 192)    ' C586C0
        Case tkKeywordDecl:  ThemeColor = RGB(86, 156, 214)     ' 569CD6
        Case tkFunction:     ThemeColor = RGB(220, 220, 170)    ' DCDCAA
        Case tkClass:        ThemeColor = RGB(78, 201, 176)     ' 4EC9B0
        Case tkVariable:     ThemeColor = RGB(156, 220, 254)    ' 9CDCFE
        Case tkBracket1:     ThemeColor = RGB(255, 215, 0)      ' FFD700 gold
        Case tkBracket2:     ThemeColor = RGB(218, 112, 214)    ' DA70D6 orchid
        Case tkBracket3:     ThemeColor = RGB(23, 159, 255)     ' 179FFF blue
        Case Else:           ThemeColor = RGB(212, 212, 212)    ' D4D4D4 default
    End Select
End Function

Public Function ThemeBackColor() As Long
    ThemeBackColor = RGB(31, 31, 31)                            ' 1F1F1F
End Function

Public Function ThemeGutterColor() As Long
    ThemeGutterColor = RGB(133, 133, 133)                       ' 858585
End Function

' The mask characters used by tools/lexref.py, so modSelfTest can dump what the
' VBA scanner decided in a format that diffs directly against the reference.
Public Function ThemeMaskChar(ByVal kind As TokenClass) As String
    Select Case kind
        Case tkComment:      ThemeMaskChar = "c"
        Case tkString:       ThemeMaskChar = "s"
        Case tkNumber:       ThemeMaskChar = "n"
        Case tkKeywordCtrl:  ThemeMaskChar = "k"
        Case tkKeywordDecl:  ThemeMaskChar = "d"
        Case tkFunction:     ThemeMaskChar = "f"
        Case tkClass:        ThemeMaskChar = "t"
        Case tkVariable:     ThemeMaskChar = "v"
        Case tkBracket1:     ThemeMaskChar = "1"
        Case tkBracket2:     ThemeMaskChar = "2"
        Case tkBracket3:     ThemeMaskChar = "3"
        Case Else:           ThemeMaskChar = "."
    End Select
End Function
