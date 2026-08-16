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
Public Function SpecGutter(ByVal size As Single, ByVal lineCount As Long) As Single
    Dim digits As Long
    digits = Len(CStr(IIf(lineCount < 1, 1, lineCount)))
    If digits < 2 Then digits = 2
    SpecGutter = Round(digits * ADVANCE * size + size * GAP_RATIO, 1)
End Function

' The block hugs its content vertically. A fixed height leaves dead dark space
' under a short snippet.
Public Function SpecHeight(ByVal size As Single, ByVal lineCount As Long) As Single
    SpecHeight = lineCount * SpecLine(size) + 2 * SpecPad(size)
End Function

' PowerPoint's rounded-rectangle adjustment is a fraction of the shorter side,
' so holding the radius visually constant means dividing by the actual height.
Public Function SpecCornerAdjust(ByVal size As Single, ByVal heightPt As Single) As Single
    Dim adj As Single
    If heightPt < 1 Then heightPt = 1
    adj = SpecRadius(size) / heightPt
    If adj > 0.5 Then adj = 0.5
    SpecCornerAdjust = adj
End Function
