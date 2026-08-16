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
        ' Brightened from VS Code's 6A9955 for projection, and the ONE
        ' deliberate divergence from Dark+ in this palette. Measured against the
        ' block background: 4.95:1 becomes 8.17:1. Comments were the weakest
        ' colour here and are the first thing a washed-out projector loses.
        '
        ' Lifted in saturation, not merely in lightness. A simply-lighter green
        ' (8EB67C, 7.17:1) drifts toward the number green B5CEA8 - separation
        ' falls from 124 to 64 - and comments start reading as numbers. This
        ' keeps 77.
        Case tkComment:      ThemeColor = RGB(135, 199, 107)     ' 87C76B
        Case tkString:       ThemeColor = RGB(206, 145, 120)    ' CE9178
        Case tkNumber:       ThemeColor = RGB(181, 206, 168)    ' B5CEA8
        Case tkKeywordCtrl:  ThemeColor = RGB(197, 134, 192)    ' C586C0
        Case tkKeywordDecl:  ThemeColor = RGB(86, 156, 214)     ' 569CD6
        Case tkFunction:     ThemeColor = RGB(220, 220, 170)    ' DCDCAA
        Case tkClass:        ThemeColor = RGB(78, 201, 176)     ' 4EC9B0
        ' Plain identifiers are WHITE, not light blue. Measured: two VS Code
        ' screenshots of real code contained zero 9CDCFE pixels. That colour
        ' comes from Pylance semantic highlighting, which is not in play here -
        ' the stock grammar leaves variables at the default colour.
        '
        ' The token class stays distinct even though it now shares the default
        ' colour. The lexer still knows what is a variable, so a future theme
        ' with semantic colouring is a one-line change here rather than a lexer
        ' change, and the test masks keep telling v apart from punctuation.
        Case tkVariable:     ThemeColor = RGB(212, 212, 212)    ' D4D4D4
        Case tkBracket1:     ThemeColor = RGB(255, 215, 0)      ' FFD700 gold
        Case tkBracket2:     ThemeColor = RGB(218, 112, 214)    ' DA70D6 orchid
        Case tkBracket3:     ThemeColor = RGB(23, 159, 255)     ' 179FFF blue
        Case Else:           ThemeColor = RGB(212, 212, 212)    ' D4D4D4 default
    End Select
End Function

Public Function ThemeBackColor() As Long
    ThemeBackColor = RGB(31, 31, 31)                            ' 1F1F1F
End Function

' The band drawn behind emphasised lines. Light enough to read as deliberate
' from the back of a room, dark enough that the code on top stays legible -
' every token colour still clears WCAG AAA against it.
Public Function ThemeEmphasisColor() As Long
    ThemeEmphasisColor = RGB(58, 68, 82)                        ' 3A4452
End Function

' A colour faded towards the block background. Used to push everything except
' the emphasised lines back, which reads far more strongly from the back of a
' room than brightening the emphasised lines alone: the eye is drawn by
' CONTRAST, and dimming the surroundings raises it everywhere at once.
Public Function ThemeDimmed(ByVal rgbColor As Long) As Long
    Const KEEP As Single = 0.4          ' how much of the original survives
    Dim r As Long, g As Long, b As Long
    Dim br As Long, bg As Long, bb As Long, back As Long

    back = ThemeBackColor()
    r = rgbColor And &HFF&
    g = (rgbColor \ &H100&) And &HFF&
    b = (rgbColor \ &H10000) And &HFF&
    br = back And &HFF&
    bg = (back \ &H100&) And &HFF&
    bb = (back \ &H10000) And &HFF&

    ThemeDimmed = RGB(br + (r - br) * KEEP, _
                      bg + (g - bg) * KEEP, _
                      bb + (b - bb) * KEEP)
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
