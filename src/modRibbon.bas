Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon - ribbon callbacks.
'
' PHASE 0 STATE: every callback named in ribbon/customUI14.xml exists here, but
' only the language dropdown does real work. The rest report themselves and
' return. That is deliberate - Phase 0 proves the packaging and the ribbon wiring
' before any real code is written, and a stub that fires is proof.
'
' Ribbon callbacks must match the signatures Office expects exactly. A wrong
' signature fails at click time with "Cannot run the macro", which looks like a
' packaging problem and is not.
'==============================================================================
Option Explicit

Private mRibbon     As IRibbonUI
Private mLangIndex  As Long        ' Phase 1: read from the selected shape's tag

Public Const ADDIN_NAME As String = "CodeHighlight"

'------------------------------------------------------------------------------
' Load
'------------------------------------------------------------------------------

' Cached so Phase 1 can call mRibbon.Invalidate when the selection changes and
' the dropdown / toggle need to re-read their state from the selected shape.
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    mLangIndex = 0
End Sub

Public Sub RefreshRibbon()
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

'------------------------------------------------------------------------------
' Language dropdown - driven entirely by modLangRegistry.
' These three callbacks never need to change when a language is added.
'------------------------------------------------------------------------------

Public Sub RibbonLangCount(control As IRibbonControl, ByRef count)
    count = modLangRegistry.LangCount()
End Sub

Public Sub RibbonLangLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = modLangRegistry.LangAt(CLng(index)).DisplayName
End Sub

Public Sub RibbonLangSelected(control As IRibbonControl, ByRef index)
    ' Phase 1: read the language tag off the selected block instead, so each
    ' block remembers its own language and the dropdown follows the selection.
    index = mLangIndex
End Sub

Public Sub RibbonLangChanged(control As IRibbonControl, id As String, index As Integer)
    mLangIndex = CLng(index)
    ' Phase 1: write the tag onto the selected block, then re-highlight it.
    Spike "Language set to " & modLangRegistry.LangAt(mLangIndex).DisplayName
End Sub

'------------------------------------------------------------------------------
' Buttons - Phase 1 onward fills these in
'------------------------------------------------------------------------------

Public Sub RibbonNewBlock(control As IRibbonControl)
    Spike "New block"                  ' Phase 1 -> modBlock.CreateBlock
End Sub

Public Sub RibbonHighlight(control As IRibbonControl)
    Spike "Highlight"                  ' Phase 1 -> modRender.ApplyHighlight
End Sub

Public Sub RibbonHighlightAll(control As IRibbonControl)
    Spike "Highlight all"              ' Phase 4
End Sub

Public Sub RibbonGutterPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = False                ' Phase 3 -> does the selected block have a gutter
End Sub

Public Sub RibbonToggleGutter(control As IRibbonControl, pressed As Boolean)
    Spike "Line numbers " & IIf(pressed, "on", "off")   ' Phase 3
End Sub

Public Sub RibbonSizeUp(control As IRibbonControl)
    Spike "Larger"                     ' Phase 4, capped at the fitting size
End Sub

Public Sub RibbonSizeDown(control As IRibbonControl)
    Spike "Smaller"                    ' Phase 4
End Sub

Public Sub RibbonFit(control As IRibbonControl)
    Spike "Fit"                        ' Phase 4, warns below MIN_TEACHING_SIZE
End Sub

'------------------------------------------------------------------------------

' The Phase 0 proof: the tab loaded, the button fired, and the registry answered.
Private Sub Spike(ByVal what As String)
    MsgBox what & " - not implemented yet." & vbCrLf & vbCrLf & _
           "Languages registered: " & modLangRegistry.LangCount() & _
           " (" & modLangRegistry.LangAt(mLangIndex).DisplayName & " selected)", _
           vbInformation, ADDIN_NAME
End Sub
