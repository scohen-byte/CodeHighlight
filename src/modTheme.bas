Attribute VB_Name = "modTheme"
'==============================================================================
' modTheme - token classes and their colours.
'
' Token classes are LANGUAGE-NEUTRAL. Comment, String, Number, Keyword,
' Function, Class and Variable are concepts every language in scope shares, so
' adding a language never touches this module. That is the whole reason the
' palette sits behind a function instead of being inlined in the renderer.
'
' There are TWO themes, and which one a block uses is a property of that block.
' See "The current theme" below for how it is threaded, and modRetheme for how
' a block is moved from one to the other.
'
'   Dark  - VS Code Dark Modern, matching tools/lab.py exactly. lab.py is the
'           spec for this theme - when this disagrees with it, this is wrong.
'   Light - GitHub Light Default, from primer/github-vscode-theme, darkened
'           only where it failed a contrast floor. See ThemeColor.
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

Public Enum ThemeId
    thDark = 0
    thLight = 1
End Enum

Public Const THEME_FONT As String = "Consolas"

'------------------------------------------------------------------------------
' The current theme
'------------------------------------------------------------------------------
' Every colour function below reads this rather than taking a theme argument.
' The alternative - a parameter on all fourteen of them - would have touched
' every one of the thirty-odd call sites and every signature, to say the same
' thing at each: "the theme of the block we are drawing".
'
' The cost is module state, and the failure mode it buys is a block rendered in
' the PREVIOUS block's theme. So there is exactly one writer: modRender.Apply-
' Highlight sets it from the shape's own tag before it draws anything, and
' modRetheme sets it around the colours it writes. Nothing else may set it.
' ThemeTest in modSelfTest renders two differently-themed blocks back to
' back and asserts the second did not inherit the first, because that is the
' bug this design can have.
Private mTheme As ThemeId

Public Function ThemeCurrent() As ThemeId
    ThemeCurrent = mTheme
End Function

Public Sub ThemeSetCurrent(ByVal t As ThemeId)
    If t <> thLight Then t = thDark
    mTheme = t
End Sub

Public Function ThemeCount() As Long
    ThemeCount = 2
End Function

Public Function ThemeName(ByVal i As Long) As String
    Select Case i
        Case thLight: ThemeName = "Light"
        Case Else:    ThemeName = "Dark"
    End Select
End Function

' Round-trips a stored tag. An UNSET tag has to mean Dark: every deck made
' before this feature existed has no theme tag on any block, and those blocks
' are dark. That single default is the whole of the backward compatibility
' story - there is no migration and no repair pass.
Public Function ThemeFromTag(ByVal v As String) As ThemeId
    If v = "1" Then ThemeFromTag = thLight Else ThemeFromTag = thDark
End Function

Public Function ThemeToTag(ByVal t As ThemeId) As String
    If t = thLight Then ThemeToTag = "1" Else ThemeToTag = "0"
End Function

'------------------------------------------------------------------------------
' Token colours
'------------------------------------------------------------------------------
Public Function ThemeColor(ByVal kind As TokenClass) As Long
    If mTheme = thLight Then
        ThemeColor = LightColor(kind)
    Else
        ThemeColor = DarkColor(kind)
    End If
End Function

Private Function DarkColor(ByVal kind As TokenClass) As Long
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
        Case tkComment:      DarkColor = RGB(135, 199, 107)     ' 87C76B
        Case tkString:       DarkColor = RGB(206, 145, 120)     ' CE9178
        Case tkNumber:       DarkColor = RGB(181, 206, 168)     ' B5CEA8
        Case tkKeywordCtrl:  DarkColor = RGB(197, 134, 192)     ' C586C0
        Case tkKeywordDecl:  DarkColor = RGB(86, 156, 214)      ' 569CD6
        Case tkFunction:     DarkColor = RGB(220, 220, 170)     ' DCDCAA
        Case tkClass:        DarkColor = RGB(78, 201, 176)      ' 4EC9B0
        ' Plain identifiers are WHITE, not light blue. Measured: two VS Code
        ' screenshots of real code contained zero 9CDCFE pixels. That colour
        ' comes from Pylance semantic highlighting, which is not in play here -
        ' the stock grammar leaves variables at the default colour.
        '
        ' The token class stays distinct even though it now shares the default
        ' colour. The lexer still knows what is a variable, so a future theme
        ' with semantic colouring is a one-line change here rather than a lexer
        ' change, and the test masks keep telling v apart from punctuation.
        Case tkVariable:     DarkColor = RGB(212, 212, 212)     ' D4D4D4
        Case tkBracket1:     DarkColor = RGB(255, 215, 0)       ' FFD700 gold
        Case tkBracket2:     DarkColor = RGB(218, 112, 214)     ' DA70D6 orchid
        Case tkBracket3:     DarkColor = RGB(23, 159, 255)      ' 179FFF blue
        Case Else:           DarkColor = RGB(212, 212, 212)     ' D4D4D4 default
    End Select
End Function

' GitHub Light Default, from primer/github-vscode-theme.
'
' Taken from a published theme rather than invented. An earlier hand-tuned
' palette scored WORSE than this one on every axis that matters: it was built by
' darkening each colour to the same contrast target, which put them all at the
' same lightness and left hue as the only thing telling them apart - and hue is
' exactly what red-green colour blindness removes. GitHub's palette has a
' natural lightness spread and was designed for accessibility. Measured, its
' worst pair under deuteranopia is 8.5 against the dark theme's 5.3.
'
' Four colours are exactly as GitHub publishes them. The other seven were too
' light for a projector - GitHub tunes for a monitor at arm's length - so each
' lost lightness along its own hue until it cleared 5.5:1 on the block ground.
' That floor is the dark theme's current worst (5.59:1), so the two themes are
' at parity in a dim hall and the light one pulls ahead as the lights come up:
' 4.36:1 against 3.36:1 with the lights on, 1.30x better.
Private Function LightColor(ByVal kind As TokenClass) As Long
    Select Case kind
        Case tkComment:      LightColor = RGB(93, 102, 111)     ' 5D666F darkened
        Case tkString:       LightColor = RGB(10, 48, 105)      ' 0A3069 as published
        Case tkNumber:       LightColor = RGB(5, 80, 174)       ' 0550AE as published
        ' GitHub does not distinguish flow control from declaration - both are
        ' keywords and both are red. Kept that way rather than inventing a
        ' second hue it never had.
        Case tkKeywordCtrl:  LightColor = RGB(199, 22, 41)      ' C71629 darkened
        Case tkKeywordDecl:  LightColor = RGB(199, 22, 41)      ' C71629 darkened
        Case tkFunction:     LightColor = RGB(118, 69, 211)     ' 7645D3 darkened
        Case tkClass:        LightColor = RGB(149, 56, 0)       ' 953800 as published
        Case tkVariable:     LightColor = RGB(31, 35, 40)       ' 1F2328 as published
        ' GitHub's own bracket highlight colours, darkened. They only have to be
        ' told apart from EACH OTHER: a bracket is a punctuation glyph, and
        ' nobody confuses "(" with an identifier, so a bracket sharing a hue
        ' with a keyword costs nothing. Measured that way the three levels stay
        ' 5.6 to 10.2 apart even for a dichromat.
        Case tkBracket1:     LightColor = RGB(0, 97, 203)       ' 0061CB darkened
        Case tkBracket2:     LightColor = RGB(4, 117, 46)       ' 04752E darkened
        Case tkBracket3:     LightColor = RGB(137, 91, 0)       ' 895B00 darkened
        Case Else:           LightColor = RGB(31, 35, 40)       ' 1F2328 default
    End Select
End Function

'------------------------------------------------------------------------------
' Surfaces
'------------------------------------------------------------------------------
' The light ground is GitHub's own canvas.subtle rather than pure white, so the
' block is a surface rather than a hole in the slide.
Public Function ThemeBackColor() As Long
    If mTheme = thLight Then
        ThemeBackColor = RGB(246, 248, 250)                     ' F6F8FA
    Else
        ThemeBackColor = RGB(31, 31, 31)                        ' 1F1F1F
    End If
End Function

' A dark block on a white slide is its own outline. A light one is not: F6F8FA
' against white is 1.06:1, which is invisible from the back of a room, so the
' border is the only thing making the block an object. GitHub's
' borderColor.default, which is what their own panels use.
Public Function ThemeBlockHasEdge() As Boolean
    ThemeBlockHasEdge = (mTheme = thLight)
End Function

Public Function ThemeBlockEdge() As Long
    ThemeBlockEdge = RGB(208, 215, 222)                         ' D0D7DE
End Function

' The band drawn behind emphasised lines. Light enough to read as deliberate
' from the back of a room, dark enough that the code on top stays legible.
Public Function ThemeEmphasisColor() As Long
    If mTheme = thLight Then
        ThemeEmphasisColor = RGB(233, 235, 238)                 ' E9EBEE
    Else
        ThemeEmphasisColor = RGB(58, 68, 82)                    ' 3A4452
    End If
End Function

' How much of a colour survives dimming. The dark theme keeps 0.4; the light
' theme needs more, because a blend toward a light ground loses contrast far
' faster - at 0.4 the weakest dimmed token measures 1.77:1, which is a texture
' rather than text. 0.6 brings it to 2.47:1.
'
' Dimmed code still has to be READ. It is the code the class has not reached
' yet, not code that has been deleted.
Private Function DimKeep() As Single
    If mTheme = thLight Then DimKeep = 0.6 Else DimKeep = 0.4
End Function

' A colour faded towards the block background. Used to push everything except
' the emphasised lines back, which reads far more strongly from the back of a
' room than brightening the emphasised lines alone: the eye is drawn by
' CONTRAST, and dimming the surroundings raises it everywhere at once.
Public Function ThemeDimmed(ByVal rgbColor As Long) As Long
    Dim r As Long, g As Long, b As Long
    Dim br As Long, bg As Long, bb As Long, back As Long
    Dim keep As Single

    keep = DimKeep()
    back = ThemeBackColor()
    r = rgbColor And &HFF&
    g = (rgbColor \ &H100&) And &HFF&
    b = (rgbColor \ &H10000) And &HFF&
    br = back And &HFF&
    bg = (back \ &H100&) And &HFF&
    bb = (back \ &H10000) And &HFF&

    ThemeDimmed = RGB(br + (r - br) * keep, _
                      bg + (g - bg) * keep, _
                      bb + (b - bb) * keep)
End Function

' The panel that hides code awaiting a reveal. Close to the block background so
' it reads as absence rather than as a new element, but not identical - a plain
' hole looks like a rendering fault, and this should look deliberate.
Public Function ThemeCoverColor() As Long
    If mTheme = thLight Then
        ThemeCoverColor = RGB(239, 241, 243)                    ' EFF1F3
    Else
        ThemeCoverColor = RGB(42, 42, 42)                       ' 2A2A2A
    End If
End Function

' The question mark on a cover panel, set bold where it is used: it is the only
' thing on that part of the slide, and it is the question being put to the room.
' Whichever end of the range the cover is not.
Public Function ThemeCoverMarkColor() As Long
    If mTheme = thLight Then
        ThemeCoverMarkColor = RGB(31, 35, 40)                   ' 1F2328
    Else
        ThemeCoverMarkColor = RGB(255, 255, 255)                ' FFFFFF
    End If
End Function

'------------------------------------------------------------------------------
' Notes
'------------------------------------------------------------------------------
' The note PRESETS are deliberately NOT theme-aware, and neither are arrows.
' Both sit on the SLIDE rather than on the block, so their readability is a
' property of their own fill and the slide behind them - not of which theme the
' code is wearing. Making them follow the theme would repaint choices the user
' made, which is the one thing a theme switch must not do.
'
' Only the DEFAULT changes: a brand new note in a light deck starts as Paper
' rather than Slate, because that is the preset that belongs beside a light
' block. An existing note keeps whatever it was given.
Public Function ThemeNoteColor() As Long
    If mTheme = thLight Then
        ThemeNoteColor = ThemeNotePreset(5)                     ' Paper F2F2EF
    Else
        ThemeNoteColor = ThemeNotePreset(0)                     ' Slate 3A4452
    End If
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
        Case 0: ThemeNoteFontName = "Default font"
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

Public Function ThemeTextOn(ByVal rgbColor As Long) As Long
    Const DARK As Long = 1710618                                ' RGB(26,26,26)
    If ThemeContrast(rgbColor, DARK) >= ThemeContrast(rgbColor, ThemeNoteTextColor()) Then
        ThemeTextOn = DARK
    Else
        ThemeTextOn = ThemeNoteTextColor()
    End If
End Function

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

'------------------------------------------------------------------------------
' Furniture
'------------------------------------------------------------------------------
' The line from a note back to its code. Quiet on purpose: it has to be
' followable without competing with either end of it.
Public Function ThemeLeaderColor() As Long
    If mTheme = thLight Then
        ThemeLeaderColor = RGB(140, 149, 159)                   ' 8C959F
    Else
        ThemeLeaderColor = RGB(133, 133, 133)                   ' 858585
    End If
End Function

' An output note: the terminal, borrowed. Set apart from the code block the way
' VS Code's panel is set apart from its editor, so it reads as a different
' surface rather than as a second block of code.
Public Function ThemeOutputFill() As Long
    If mTheme = thLight Then
        ThemeOutputFill = RGB(237, 239, 241)                    ' EDEFF1
    Else
        ThemeOutputFill = RGB(24, 24, 24)                       ' 181818
    End If
End Function

Public Function ThemeOutputText() As Long
    If mTheme = thLight Then
        ThemeOutputText = RGB(31, 35, 40)                       ' 1F2328
    Else
        ThemeOutputText = RGB(204, 204, 204)                    ' CCCCCC
    End If
End Function

' Two surfaces of the same weight on one slide need an edge, or they merge into
' a single silhouette from the back of a room.
Public Function ThemeOutputEdge() As Long
    If mTheme = thLight Then
        ThemeOutputEdge = RGB(175, 184, 193)                    ' AFB8C1
    Else
        ThemeOutputEdge = RGB(90, 90, 90)                       ' 5A5A5A
    End If
End Function

' The arrow that opens an output line. In the dark theme this is the comment
' green, which is the palette's one colour that already means "not the code" -
' but it is stated outright rather than read from ThemeColor(tkComment), so that
' retuning comments for a projector does not silently move it.
'
' In the light theme the two genuinely part company: GitHub's comment is a grey,
' which would say nothing at all here, so the mark takes the bracket green
' instead and keeps meaning what it meant.
Public Function ThemeOutputMark() As Long
    If mTheme = thLight Then
        ThemeOutputMark = RGB(4, 117, 46)                       ' 04752E
    Else
        ThemeOutputMark = RGB(135, 199, 107)                    ' 87C76B
    End If
End Function

Public Function ThemeGutterColor() As Long
    If mTheme = thLight Then
        ThemeGutterColor = RGB(110, 119, 129)                   ' 6E7781
    Else
        ThemeGutterColor = RGB(133, 133, 133)                   ' 858585
    End If
End Function

'------------------------------------------------------------------------------
' Arrows
'------------------------------------------------------------------------------
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
' and what a code deck almost always uses - which is also why it does NOT vary
' by theme. An arrow never sits on the block, so the block's ground is not the
' background it has to survive.
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
