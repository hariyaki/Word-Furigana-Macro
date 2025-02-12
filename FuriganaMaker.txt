Option Explicit

' If Word doesn稚 have this constant, define it manually:
Const wdDialogFormatPhoneticGuide As Long = 986

'---------------------------------------------------------------------------------
' SINGLE PUBLIC MACRO:
'   1) Unlinks all fields (so no "EQ \*..." expansions),
'   2) Applies furigana to consecutive Kanji runs from doc end ? doc start.
'
' This is the ONLY macro that appears in the Macros menu.
'---------------------------------------------------------------------------------
Public Sub FuriganaMaker()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' 1) Unlink all fields (removes them, but preserves other formatting)
    doc.Fields.Unlink
    
    ' 2) Ensure doc is recognized as Japanese
    doc.Content.LanguageIDFarEast = wdJapanese
    
    ' 3) Call our private sub to scan backwards for Kanji
    BackwardLoop doc
End Sub

'---------------------------------------------------------------------------------
' PRIVATE SUB: Scans doc from end to start, finds consecutive Kanji runs,
' then applies the Phonetic Guide dialog to each run via hacky SendKeys.
'
' Marked Private, so it won't appear in the macro list.
'---------------------------------------------------------------------------------
Private Sub BackwardLoop(ByVal doc As Document)
    Dim docRange As Range
    Dim docStart As Long, docEnd As Long
    Dim i As Long
    Dim inRun As Boolean
    Dim runEnd As Long
    Dim runCount As Long
    
    Set docRange = doc.Content
    docStart = docRange.Start
    docEnd = docRange.End
    
    inRun = False
    runCount = 0
    
    ' Loop backwards
    For i = (docEnd - 1) To docStart Step -1
        Dim ch As String
        ch = doc.Range(i, i + 1).Text
        
        If IsKanjiChar(ch) Then
            If Not inRun Then
                inRun = True
                runEnd = i
            End If
        Else
            If inRun Then
                ApplyPhoneticGuideToRange (i + 1), runEnd
                runCount = runCount + 1
                inRun = False
            End If
        End If
    Next i
    
    ' If still in a run at docStart
    If inRun Then
        If runEnd >= docStart Then
            ApplyPhoneticGuideToRange docStart, runEnd
            runCount = runCount + 1
        End If
        inRun = False
    End If
    
    MsgBox "Furigana applied to " & runCount & " Kanji runs (backwards).", vbInformation
End Sub

'---------------------------------------------------------------------------------
' PRIVATE SUB: Applies Word's Phonetic Guide to [startPos..endPos].
' Uses SendKeys "~" to auto-press "OK" in the Phonetic Guide dialog.
'---------------------------------------------------------------------------------
Private Sub ApplyPhoneticGuideToRange(startPos As Long, endPos As Long)
    If endPos < startPos Then Exit Sub
    
    Dim rng As Range
    Dim WAIT_SECONDS As Single
    WAIT_SECONDS = 0.5
    
    Set rng = ActiveDocument.Range(startPos, endPos + 1)
    rng.LanguageIDFarEast = wdJapanese
    rng.Select
    
    WaitSeconds WAIT_SECONDS
    
    ' Queue the Enter key
    SendKeys "~", False
    
    On Error Resume Next
    Dialogs(wdDialogFormatPhoneticGuide).Show
    On Error GoTo 0
    
    WaitSeconds WAIT_SECONDS
End Sub

'---------------------------------------------------------------------------------
' PRIVATE FUNCTION: Checks if 'ch' is a single character in standard/Ext-A Kanji,
' handling negative AscW values.
'---------------------------------------------------------------------------------
Private Function IsKanjiChar(ch As String) As Boolean
    Dim code As Long
    IsKanjiChar = False
    
    If Len(ch) = 1 Then
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        
        ' Standard CJK (U+4E00..U+9FFF) => decimal 19968..40959
        If code >= 19968 And code <= 40959 Then
            IsKanjiChar = True
            
        ' Extension A (U+3400..U+4DBF)
        ElseIf code >= &H3400 And code <= &H4DBF Then
            IsKanjiChar = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------
' PRIVATE SUB: Simple wait to let Word open/apply the dialog
'---------------------------------------------------------------------------------
Private Sub WaitSeconds(ByVal Seconds As Single)
    Dim t As Date
    t = DateAdd("s", Seconds, Now)
    Do While Now < t
        DoEvents
    Loop
End Sub
