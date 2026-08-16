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

' The panel that hides code awaiting a reveal. Close to the block background so
' it reads as absence rather than as a new element, but not identical - a plain
' hole looks like a rendering fault, and this should look deliberate.
Public Function ThemeCoverColor() As Long
    ThemeCoverColor = RGB(42, 42, 42)                           ' 2A2A2A
End Function

' A note attached to a line. The same slate as the emphasis band, at full
' opacity: a note is an aside about the code, so it should read as part of the
' same object rather than as a sticker dropped on the slide. Against it, the
' note text below clears 9.8:1.
Public Function ThemeNoteColor() As Long
    ThemeNoteColor = RGB(58, 68, 82)                            ' 3A4452
End Function

' Brighter than the code default. A note is prose read from the back of a room,
' and it is set smaller than the code, so it gets the contrast back.
Public Function ThemeNoteTextColor() As Long
    ThemeNoteTextColor = RGB(235, 235, 235)                     ' EBEBEB
End Function

' The colours a note can be set to. A fixed list rather than a full colour
' picker, because every one of these is checked against the text colour it gets
' and an arbitrary colour is not - OptionsTest asserts each clears 4.5:1.
'
' Two halves. The first six are quiet, and belong beside a dark code block: they
' read as an aside about the code. The rest are saturated, for a note that is
' meant to be the loudest thing on the slide - a warning, a gotcha, the answer.
' The quiet six alone turned out to be too close to tell apart in the gallery.
'
' Slate is first because it is the default.
Public Function ThemeNotePresetCount() As Long
    ThemeNotePresetCount = 12
End Function

Public Function ThemeNotePresetName(ByVal i As Long) As String
    Select Case i
        Case 0: ThemeNotePresetName = "Slate"
        Case 1: ThemeNotePresetName = "Indigo"
        Case 2: ThemeNotePresetName = "Teal"
        Case 3: ThemeNotePresetName = "Plum"
        Case 4: ThemeNotePresetName = "Ink"
        Case 5: ThemeNotePresetName = "Paper"
        Case 6: ThemeNotePresetName = "Crimson"
        Case 7: ThemeNotePresetName = "Rose"
        Case 8: ThemeNotePresetName = "Amber"
        Case 9: ThemeNotePresetName = "Emerald"
        Case 10: ThemeNotePresetName = "Royal"
        Case Else: ThemeNotePresetName = "Violet"
    End Select
End Function

Public Function ThemeNotePreset(ByVal i As Long) As Long
    Select Case i
        Case 0: ThemeNotePreset = RGB(58, 68, 82)               ' 3A4452
        Case 1: ThemeNotePreset = RGB(59, 58, 92)               ' 3B3A5C
        Case 2: ThemeNotePreset = RGB(30, 70, 69)               ' 1E4645
        Case 3: ThemeNotePreset = RGB(74, 51, 80)               ' 4A3350
        Case 4: ThemeNotePreset = RGB(43, 43, 43)               ' 2B2B2B
        Case 5: ThemeNotePreset = RGB(242, 242, 239)            ' F2F2EF
        Case 6: ThemeNotePreset = RGB(198, 40, 40)              ' C62828
        Case 7: ThemeNotePreset = RGB(194, 24, 91)              ' C2185B
        ' Light enough to take dark text, which is what makes it usable at all:
        ' a saturated amber with white words on it fails contrast outright.
        Case 8: ThemeNotePreset = RGB(255, 179, 0)              ' FFB300
        Case 9: ThemeNotePreset = RGB(0, 121, 80)               ' 007950
        Case 10: ThemeNotePreset = RGB(29, 78, 216)             ' 1D4ED8
        Case Else: ThemeNotePreset = RGB(109, 40, 217)          ' 6D28D9
    End Select
End Function

' WCAG relative luminance, with the sRGB gamma curve. Used to check that every
' preset above is actually readable with the text colour it is given, rather
' than to make that claim and hope.
Public Function ThemeLuminance(ByVal rgbColor As Long) As Double
    Dim r As Long, g As Long, b As Long
    r = rgbColor And &HFF&
    g = (rgbColor \ &H100&) And &HFF&
    b = (rgbColor \ &H10000) And &HFF&
    ThemeLuminance = 0.2126 * Linearize(r) + 0.7152 * Linearize(g) + 0.0722 * Linearize(b)
End Function

Private Function Linearize(ByVal channel As Long) As Double
    Dim c As Double
    c = channel / 255
    If c <= 0.03928 Then
        Linearize = c / 12.92
    Else
        Linearize = ((c + 0.055) / 1.055) ^ 2.4
    End If
End Function

Public Function ThemeContrast(ByVal a As Long, ByVal b As Long) As Double
    Dim la As Double, lb As Double, hi As Double, lo As Double
    la = ThemeLuminance(a)
    lb = ThemeLuminance(b)
    If la > lb Then
        hi = la: lo = lb
    Else
        hi = lb: lo = la
    End If
    ThemeContrast = (hi + 0.05) / (lo + 0.05)
End Function

' The fonts a note can be set to. Item 0 is the deck's own body font, which is
' the default: a note is prose, and inheriting the theme font makes it look like
' the rest of the presentation. The rest ship with Windows or with Office, so
' nothing here can fail to render on a machine that can open the deck.
Public Function ThemeNoteFontCount() As Long
    ThemeNoteFontCount = 5
End Function

Public Function ThemeNoteFontName(ByVal i As Long) As String
    Select Case i
        Case 0: ThemeNoteFontName = "Deck default"
        Case 1: ThemeNoteFontName = "Segoe UI"
        Case 2: ThemeNoteFontName = "Calibri"
        Case 3: ThemeNoteFontName = "Georgia"
        Case Else: ThemeNoteFontName = "Consolas"
    End Select
End Function

' The value to hand to Font.Name. Empty for the deck default, which means
' "leave it alone".
Public Function ThemeNoteFontValue(ByVal i As Long) As String
    If i = 0 Then Exit Function
    ThemeNoteFontValue = ThemeNoteFontName(i)
End Function

Public Function ThemeNoteFontIndexOf(ByVal fontName As String) As Long
    Dim i As Long
    If Len(fontName) = 0 Then Exit Function
    For i = 1 To ThemeNoteFontCount() - 1
        If ThemeNoteFontName(i) = fontName Then
            ThemeNoteFontIndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function ThemeNotePresetIndexOf(ByVal rgbColor As Long) As Long
    Dim i As Long
    For i = 0 To ThemeNotePresetCount() - 1
        If ThemeNotePreset(i) = rgbColor Then
            ThemeNotePresetIndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function ThemeIsLight(ByVal rgbColor As Long) As Boolean
    ThemeIsLight = (ThemeLuminance(rgbColor) > 0.5)
End Function

' The text colour for a given note fill: whichever of the two reads better on
' it. Picking by MEASURED contrast rather than by a lightness threshold is what
' makes a saturated preset safe - amber is light enough to need dark words and
' crimson is not, and no single cutoff gets both right by luck.
Public Function ThemeTextOn(ByVal rgbColor As Long) As Long
    Const DARK As Long = 1710618                                ' RGB(26,26,26)
    If ThemeContrast(rgbColor, DARK) >= ThemeContrast(rgbColor, ThemeNoteTextColor()) Then
        ThemeTextOn = DARK
    Else
        ThemeTextOn = ThemeNoteTextColor()
    End If
End Function

' A light note on a white slide has no edge of its own, and reads as text
' floating in space rather than as a note. Dark fills need no such help, so they
' get none - a border on them only adds a line to look at.
Public Function ThemeNeedsEdge(ByVal rgbColor As Long) As Boolean
    ThemeNeedsEdge = ThemeIsLight(rgbColor)
End Function

Public Function ThemeEdgeFor(ByVal rgbColor As Long) As Long
    Const DARKEN As Single = 0.78
    Dim r As Long, g As Long, b As Long
    r = rgbColor And &HFF&
    g = (rgbColor \ &H100&) And &HFF&
    b = (rgbColor \ &H10000) And &HFF&
    ThemeEdgeFor = RGB(r * DARKEN, g * DARKEN, b * DARKEN)
End Function

' The line from a note back to its code. Quiet on purpose: it has to be
' followable without competing with either end of it.
Public Function ThemeLeaderColor() As Long
    ThemeLeaderColor = RGB(133, 133, 133)                       ' 858585
End Function

' The question mark on a cover panel. Pure white, and set bold where it is used:
' it is the only thing on that part of the slide, and it is the question being
' put to the room.
Public Function ThemeCoverMarkColor() As Long
    ThemeCoverMarkColor = RGB(255, 255, 255)                    ' FFFFFF
End Function

' The arrow in the left margin, and the colours it can be set to.
'
' These are the SYNTAX HUES, DARKENED. Taking the palette colours as they stand
' does not work: they are chosen against the block's 1F1F1F, and an arrow sits
' on the slide, which is white. Measured against white, the teal of a class name
' manages 2.0:1 and the purple of a keyword 2.8:1 - as text on the block they
' are right, as a solid shape on a white slide they look faded, which is the
' opposite of what an arrow is for.
'
' So each keeps its hue and loses enough lightness to clear 4.5:1 on white,
' which ArrowTest asserts. The result still reads as belonging to the same
' design without pretending the background has not changed.
'
' The list is aimed at a light slide, because that is what PowerPoint gives you
' and what a code deck almost always uses.
Public Function ThemeArrowPresetCount() As Long
    ThemeArrowPresetCount = 8
End Function

Public Function ThemeArrowPresetName(ByVal i As Long) As String
    Select Case i
        Case 0: ThemeArrowPresetName = "Blue"
        Case 1: ThemeArrowPresetName = "Purple"
        Case 2: ThemeArrowPresetName = "Teal"
        Case 3: ThemeArrowPresetName = "Green"
        Case 4: ThemeArrowPresetName = "Orange"
        Case 5: ThemeArrowPresetName = "Gold"
        Case 6: ThemeArrowPresetName = "Crimson"
        Case Else: ThemeArrowPresetName = "Ink"
    End Select
End Function

Public Function ThemeArrowPreset(ByVal i As Long) As Long
    Select Case i
        ' from 569CD6, the declaration keyword blue
        Case 0: ThemeArrowPreset = RGB(31, 111, 178)            ' 1F6FB2
        ' from C586C0, the flow-control purple
        Case 1: ThemeArrowPreset = RGB(142, 76, 138)            ' 8E4C8A
        ' from 4EC9B0, the class teal
        Case 2: ThemeArrowPreset = RGB(15, 122, 104)            ' 0F7A68
        ' from 87C76B, the comment green
        Case 3: ThemeArrowPreset = RGB(66, 118, 47)             ' 42762F
        ' from CE9178, the string orange
        Case 4: ThemeArrowPreset = RGB(168, 90, 56)             ' A85A38
        ' from DCDCAA and the gold bracket
        Case 5: ThemeArrowPreset = RGB(138, 109, 0)             ' 8A6D00
        ' Not from the palette. For the line that is the mistake.
        Case 6: ThemeArrowPreset = RGB(198, 40, 40)             ' C62828
        Case Else: ThemeArrowPreset = RGB(43, 43, 43)           ' 2B2B2B
    End Select
End Function

Public Function ThemeArrowPresetIndexOf(ByVal rgbColor As Long) As Long
    Dim i As Long
    For i = 0 To ThemeArrowPresetCount() - 1
        If ThemeArrowPreset(i) = rgbColor Then
            ThemeArrowPresetIndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function ThemeArrowColor() As Long
    ThemeArrowColor = ThemeArrowPreset(0)
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
