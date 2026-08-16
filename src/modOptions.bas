Attribute VB_Name = "modOptions"
'==============================================================================
' modOptions - the choices that belong to a DECK rather than to a shape.
'
' Note size, note colour, note font, and the two walkthrough options. All of
' them are stored as tags on the Presentation, not in module variables, for two
' reasons: a module variable is lost the moment PowerPoint closes, and it would
' not travel with the file. Reopening a deck on another machine and pressing
' Stylize should not silently restyle every note.
'
' Values are stored as WHOLE NUMBERS or plain strings. CStr on a Single writes
' the locale decimal separator, so "12.5" becomes "12,5" on a German machine and
' anything parsing it gets a different number. Nothing here needs a fraction.
'
' This module exists so that modRender can read a setting without reaching into
' modRibbon. modRibbon needs IRibbonUI, which the render harness deliberately
' does not import, so a dependency that way round would make the renderer
' untestable outside a full add-in.
'==============================================================================
Option Explicit

Public Const TAG_NOTE_SIZE  As String = "CODEBLOCK_NOTE_SIZE"
Public Const TAG_NOTE_COLOR As String = "CODEBLOCK_NOTE_COLOR"
Public Const TAG_NOTE_FONT  As String = "CODEBLOCK_NOTE_FONT"
Public Const TAG_STEP_NOTE  As String = "CODEBLOCK_STEP_NOTE"
Public Const TAG_STEP_BOLD  As String = "CODEBLOCK_STEP_BOLD"

' Size 0 means "work it out from the block". An explicit size pins it.
Public Const NOTE_SIZE_AUTO As Long = 0

' What a deck that has never said otherwise gets. A fixed 24 rather than Auto,
' because Auto derives a note from the CODE size and code is set in Consolas -
' a note in a proportional face at the same nominal size looks markedly smaller,
' so deriving it produced notes that read as an afterthought. Auto is still
' there for anyone who wants notes to track the block.
Public Const NOTE_SIZE_DEFAULT As Long = 24

Private Function Pres() As Presentation
    On Error Resume Next
    Set Pres = Application.ActivePresentation
    On Error GoTo 0
End Function

Private Function GetTag(ByVal name As String) As String
    Dim p As Presentation
    Set p = Pres()
    If p Is Nothing Then Exit Function
    GetTag = p.Tags(name)
End Function

Private Sub SetTag(ByVal name As String, ByVal value As String)
    Dim p As Presentation
    Set p = Pres()
    If p Is Nothing Then Exit Sub
    p.Tags.Add name, value              ' Add replaces an existing tag
End Sub

'------------------------------------------------------------------------------
' Notes
'------------------------------------------------------------------------------

' An UNSET tag and an explicit Auto are different things, and only the string
' length tells them apart - Val("") and Val("0") are both 0. Unset means nobody
' has chosen, so it gets the default; "0" means somebody chose Auto.
Public Function NoteSize() As Long
    Dim v As String
    v = GetTag(TAG_NOTE_SIZE)
    If Len(v) = 0 Then
        NoteSize = NOTE_SIZE_DEFAULT
    Else
        NoteSize = CLng(Val(v))
    End If
End Function

Public Sub SetNoteSize(ByVal pts As Long)
    If pts < 0 Then pts = 0
    SetTag TAG_NOTE_SIZE, CStr(pts)
End Sub

Public Function NoteColor() As Long
    Dim v As String
    v = GetTag(TAG_NOTE_COLOR)
    If Len(v) = 0 Then
        NoteColor = ThemeNoteColor()
    Else
        NoteColor = CLng(Val(v))
    End If
End Function

Public Sub SetNoteColor(ByVal rgbColor As Long)
    SetTag TAG_NOTE_COLOR, CStr(rgbColor)
End Sub

' Empty means the deck's own body font, which is the default: a note is prose,
' not code, and inheriting the theme font makes it look like the rest of the
' presentation rather than like this add-in.
Public Function NoteFont() As String
    NoteFont = GetTag(TAG_NOTE_FONT)
End Function

Public Sub SetNoteFont(ByVal fontName As String)
    SetTag TAG_NOTE_FONT, fontName
End Sub

'------------------------------------------------------------------------------
' Walkthroughs
'------------------------------------------------------------------------------

' Give every generated slide a note already attached to the line it emphasises,
' so a walkthrough arrives ready to be written into rather than ready to be
' clicked at.
Public Function StepNote() As Boolean
    StepNote = (GetTag(TAG_STEP_NOTE) = "1")
End Function

Public Sub SetStepNote(ByVal on_ As Boolean)
    SetTag TAG_STEP_NOTE, IIf(on_, "1", "0")
End Sub

' Render the newly emphasised line in bold as well as banded. Applies wherever
' emphasis is used, not only in a generated walkthrough - it is a property of
' how emphasis looks, and having it differ between a hand-set block and a
' generated one would be a bug rather than a feature.
Public Function EmphasisBold() As Boolean
    EmphasisBold = (GetTag(TAG_STEP_BOLD) = "1")
End Function

Public Sub SetEmphasisBold(ByVal on_ As Boolean)
    SetTag TAG_STEP_BOLD, IIf(on_, "1", "0")
End Sub
