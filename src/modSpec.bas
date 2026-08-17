Attribute VB_Name = "modSpec"
'==============================================================================
' modSpec - the locked visual spec, as formulas.
'
' Font size is the ONLY free parameter. Everything else derives from it, so a
' block can be resized on any slide without its proportions breaking.
'
' These are the same formulas as derive() in tools/lab.py, which is the
' reference implementation. When this disagrees with lab.py, this is wrong.
' PLAN.md section 5a has the prose version.
'
' Everything here is in POINTS. PowerPoint's object model is in points, and
' mixing in inches or EMU is a good way to produce a block that is off by a
' factor of 72 and looks like a formatting bug.
'==============================================================================
Option Explicit

Public Const BASE_SIZE        As Single = 22    ' default pt size for a new block
Public Const MIN_TEACHING_SIZE As Single = 16   ' below this, warn - will not read

Private Const LINE_RATIO   As Single = 1.2     ' line spacing, EXACT points
Private Const PAD_RATIO    As Single = 0.64    ' all four internal margins
Private Const GAP_RATIO    As Single = 0.45    ' gutter-to-code gap
Private Const RADIUS_RATIO As Single = 0.36    ' corner radius
Private Const ADVANCE      As Single = 0.55    ' Consolas, from consola.ttf

' A tab advances four characters, as in VS Code. PowerPoint's own default tab
' stop is an inch - about six characters at 22pt - which is why a tab-indented
' line lands much further right than it does in an editor.
Public Const TAB_CHARS As Single = 4

' The content area on a 16:9 slide, measured in Phase -1. 960 x 540 points.
Public Const SLIDE_W  As Single = 960
Public Const SLIDE_H  As Single = 540
Public Const CONTENT_L As Single = 39.6        ' 0.55in
Public Const CONTENT_T As Single = 75.6        ' 1.05in
Public Const CONTENT_W As Single = 880.56      ' 12.23in
Public Const CONTENT_H As Single = 428.4       ' 5.95in

'------------------------------------------------------------------------------
' The size ladder
'------------------------------------------------------------------------------
' VBA has no array constants, so the ladder is a function. Kept in step with
' SIZE_LADDER in tools/lab.py.
Public Function LadderCount() As Long
    LadderCount = 10
End Function

Public Function LadderAt(ByVal i As Long) As Single
    Select Case i
        Case 0: LadderAt = 10
        Case 1: LadderAt = 12
        Case 2: LadderAt = 14
        Case 3: LadderAt = 16
        Case 4: LadderAt = 18
        Case 5: LadderAt = 20
        Case 6: LadderAt = 22
        Case 7: LadderAt = 24
        Case 8: LadderAt = 28
        Case Else: LadderAt = 32
    End Select
End Function

' Nearest rung at or below the given size, so a block set to some off-ladder
' size still steps sensibly rather than jumping.
Public Function LadderIndexOf(ByVal size As Single) As Long
    Dim i As Long, best As Long
    For i = 0 To LadderCount() - 1
        If LadderAt(i) <= size + 0.01 Then best = i
    Next i
    LadderIndexOf = best
End Function

' The width a block needs at a given size, and the height.
Public Function SpecWidthFor(ByVal size As Single, ByVal maxChars As Long, _
                             ByVal gutterW As Single) As Single
    SpecWidthFor = maxChars * SpecCharW(size) + gutterW + 2 * SpecPad(size)
End Function

' Largest rung at which the code fits the content area, in BOTH directions.
' Drives Fit, and also caps Larger so growing a block cannot push code off the
' slide - the failure the lab showed at 32pt.
Public Function SpecFitSize(ByVal lineCount As Long, ByVal maxChars As Long, _
                            ByVal withGutter As Boolean) As Single
    Dim i As Long, size As Single, gutterW As Single, best As Single

    best = LadderAt(0)
    For i = 0 To LadderCount() - 1
        size = LadderAt(i)
        gutterW = 0
        ' lineCount stands in for the highest number here, which is exact for a
        ' block numbered from 1 and slightly narrow for one numbered from 98 -
        ' a block's start number is not knowable from this signature. Worth at
        ' most one digit of width, so Fit can pick a size a shade too large on
        ' a continuation block. Say so rather than thread the block through.
        If withGutter Then gutterW = SpecGutter(size, lineCount)
        If SpecWidthFor(size, maxChars, gutterW) <= CONTENT_W Then
            If SpecHeight(size, lineCount) <= CONTENT_H Then best = size
        End If
    Next i
    SpecFitSize = best
End Function

' How many slides this code needs to stay readable, if the fitting size is
' below the teaching floor. Returns 1 when it already fits legibly.
Public Function SpecSlidesNeeded(ByVal lineCount As Long, ByVal maxChars As Long, _
                                 ByVal withGutter As Boolean) As Long
    Dim perSlide As Long

    If SpecFitSize(lineCount, maxChars, withGutter) >= MIN_TEACHING_SIZE Then
        SpecSlidesNeeded = 1
        Exit Function
    End If

    perSlide = Int((CONTENT_H - 2 * SpecPad(MIN_TEACHING_SIZE)) / SpecLine(MIN_TEACHING_SIZE))
    If perSlide < 1 Then perSlide = 1
    SpecSlidesNeeded = -Int(-lineCount / perSlide)      ' ceiling division
End Function

' Line spacing is set in EXACT points, never as a multiple. This is the single
' biggest simplification available to the gutter code later: it makes alignment
' deterministic instead of dependent on the font's internal line metrics.
Public Function SpecLine(ByVal size As Single) As Single
    SpecLine = Round(size * LINE_RATIO, 1)
End Function

Public Function SpecPad(ByVal size As Single) As Single
    SpecPad = Round(size * PAD_RATIO, 1)
End Function

Public Function SpecRadius(ByVal size As Single) As Single
    SpecRadius = Round(size * RADIUS_RATIO, 1)
End Function

Public Function SpecCharW(ByVal size As Single) As Single
    SpecCharW = size * ADVANCE
End Function

Public Function SpecTabStop(ByVal size As Single) As Single
    SpecTabStop = TAB_CHARS * SpecCharW(size)
End Function

' Width of the line-number gutter, which sits INSIDE the dark block. Not used by
' the thin slice, but it is part of the spec and belongs with the rest of it.
' Takes the HIGHEST LINE NUMBER that will be shown, which is not the line count
' when the block starts its numbering somewhere other than 1.
Public Function SpecGutter(ByVal size As Single, ByVal highestNumber As Long) As Single
    Dim digits As Long
    digits = Len(CStr(IIf(highestNumber < 1, 1, highestNumber)))
    If digits < 2 Then digits = 2
    SpecGutter = Round(digits * ADVANCE * size + size * GAP_RATIO, 1)
End Function

' The block hugs its content vertically. A fixed height leaves dead dark space
' under a short snippet.
Public Function SpecHeight(ByVal size As Single, ByVal lineCount As Long) As Single
    SpecHeight = lineCount * SpecLine(size) + 2 * SpecPad(size)
End Function

' PowerPoint's rounded-rectangle adjustment is a fraction of the SHORTER side,
' so holding the radius visually constant means dividing by whichever of width
' and height is smaller. A block that hugs its code can be narrower than it is
' tall, which is why this cannot just use the height.
Public Function SpecCornerAdjust(ByVal size As Single, ByVal shorterPt As Single) As Single
    Dim adj As Single
    If shorterPt < 1 Then shorterPt = 1
    adj = SpecRadius(size) / shorterPt
    If adj > 0.5 Then adj = 0.5
    SpecCornerAdjust = adj
End Function
