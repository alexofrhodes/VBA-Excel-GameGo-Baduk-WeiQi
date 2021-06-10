Attribute VB_Name = "GOBAN"
Global handicap As String

Sub ResetConfirm()
    Application.ScreenUpdating = True
    Dim answer As Integer
    Dim confirmation As Integer
'   confirm the reset of the board
    confirmation = MsgBox("Start a new game?", vbQuestion + vbYesNo + vbDefaultButton1, "File saved")
    If confirmation = vbNo Then
        Application.ScreenUpdating = True
        Exit Sub
    ElseIf confirmation = vbYes Then
        Application.ScreenUpdating = False
        GoReset
    End If
    Application.ScreenUpdating = True
End Sub

Sub capBtn()
[WHATCAP] = ""
handicap = ""
HandicappedGame
End Sub

Sub HandicappedGame()
    If handicap = "" Then
        handicap = InputBox("Input strength difference.")
        If handicap = "" Then Exit Sub
    End If
    Application.ScreenUpdating = False
    [GoLoop] = "Loop"
    [ksize] = 19
    GobanResize
    GoReset
'    If tmpcap <> "" Then handicap = tmpcap
    [WHATCAP] = handicap
    [komi] = 0.5
    If [Goturn] = "W" Then GoSwitch
    [GoMode] = "Setup"
    HandicapSetup
    If [GoLoop] = "" Then handicap = ""
    [gLoaded] = ""
    [pLoaded] = ""
End Sub

Sub HandicapSetup()
    Select Case handicap
    Case Is = 1
        [GoMode] = "Game"
        Application.ScreenUpdating = True
'        MsgBox "Black to play."
        Exit Sub
    Case Is = 2
        For Each cell In Range("$Q$5,$E$17")
            cell.Select
        Next cell
    Case Is = 3
        For Each cell In Range("$Q$5,$E$17,$K$11")
            cell.Select
        Next cell
    Case Is = 4
        For Each cell In Range("$Q$5,$E$17,$Q$17,$E$5")
            cell.Select
        Next cell
    Case Is = 5
        For Each cell In Range("$Q$5,$E$17,$Q$17,$E$5,$K$11")
            cell.Select
        Next cell
    Case Is = 6
        For Each cell In Range("$Q$5,$E$17,$Q$17,$E$5,$E$11,$Q$11")
            cell.Select
        Next cell
    Case Is = 7
        For Each cell In Range("$Q$5,$E$17,$Q$17,$E$5,$E$11,$Q$11,$K$11")
            cell.Select
        Next cell
    Case Is = 8
        For Each cell In Range("$Q$5,$E$17,$Q$17,$E$5,$E$11,$Q$11,$K$5,$K$17")
            cell.Select
        Next cell
    Case Is = 9
        For Each cell In Range("$Q$5,$E$17,$Q$17,$E$5,$E$11,$Q$11,$K$5,$K$17,$K$11")
            cell.Select
        Next cell
    Case Else
        MsgBox "You can only choose from 1 to 9."
'        GoReset
        [GoLoop] = ""
        Exit Sub
    End Select
    Cells(1, ActiveCell.Column).Select
    [GoMode] = "Game"
    GoSwitch
    [GoLoop] = ""
    Application.ScreenUpdating = True
'    MsgBox "White to play."
End Sub

Sub GoReset()
    If Application.ScreenUpdating = True Then Application.ScreenUpdating = False
''   save workbook to avoid loss of input
'    ThisWorkbook.Save
    
'   clear board content (letters and shapes)
    Range("Goban").ClearContents
    For Each S In ActiveSheet.Shapes
        If Not Intersect(Range(S.TopLeftCell.Address), Range("Goban")) Is Nothing Then S.Delete '  s.Select Replace:=False
    Next S
    If [komi] = 0.5 Then [komi] = 6.5
    [ScoreBlack].ClearContents
    [ScoreWhite].ClearContents
    [CountMoveBlack] = -1
    [CountMoveWhite] = -1
    [GoMovesBlack].ClearContents
    [GoMovesWhite].ClearContents
    [CapturedBlack].ClearContents
    [CapturedWhite].ClearContents
    Range("Goban").Value = 0
    [GoOperation].ClearContents
    [gLoaded] = ""
    [pLoaded] = ""
    [WHATCAP] = ""
    Cells(1, ActiveCell.Column).Select
    [GoMode] = "Game"
    [Goturn] = "B"
    ActiveSheet.Shapes("GoWhiteTurn").Visible = False
    ActiveSheet.Shapes("GoBlackTurn").Visible = True
    If [GoLoop] = "" Then Application.ScreenUpdating = True
End Sub

Sub GobanResize()
    If Application.ScreenUpdating = True Then Application.ScreenUpdating = False
    [GoLoop] = "Loop"
    GoReset
    On Error GoTo nxt
    Select Case ActiveSheet.Shapes(Application.Caller).TextFrame.Characters.Text
    Case Is = "9", "13", "19"
        [ksize] = ActiveSheet.Shapes(Application.Caller).TextFrame.Characters.Text
    End Select
nxt:
    Select Case [ksize]
    Case Is = "9"
        Range("Goban").Clear
        Range("Goban").Resize(19, 19).EntireColumn.Hidden = True
        Range("Goban").Resize(9, 9).EntireColumn.Hidden = False
        Range("Goban").Resize(9, 9).Name = "Goban"
        [fGoban].Copy
        Range("Goban").PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
        Cells(1, ActiveCell.Column).Select
        Application.CutCopyMode = False
    Case Is = "13"
        Range("Goban").Clear
        Range("Goban").Resize(19, 19).EntireColumn.Hidden = True
        Range("Goban").Resize(13, 13).EntireColumn.Hidden = False
        Range("Goban").Resize(13, 13).Name = "Goban"
        [fGoban].Copy
        Range("Goban").PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
        Cells(1, ActiveCell.Column).Select
        Application.CutCopyMode = False
        [fStars].Copy
        Range("$E$5,$K$5,$K$11,$E$11,H8").PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
        Cells(1, ActiveCell.Column).Select
        Application.CutCopyMode = False
    Case Is = "19"
        Range("Goban").Clear
        Range("Goban").Resize(19, 19).EntireColumn.Hidden = False
'        resize Goban board Named Range
        Range("Goban").Resize(19, 19).Name = "Goban"
'        Apply Goban board formatting
        [fGoban].Copy
        Range("Goban").PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
        Cells(1, ActiveCell.Column).Select
        Application.CutCopyMode = False
'        Mark the star points
        [fStars].Copy
        Range("$E$5,$K$5,$Q$5,$Q$11,$Q$17,$K$17,$K$11,$E$11,$E$17").PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
        Cells(1, ActiveCell.Column).Select
        Application.CutCopyMode = False
    End Select
    If [GoLoop] = "" Then Application.ScreenUpdating = True
End Sub

Sub Grid_Squares()
    With Cells(1, 1)
        Cells.RowHeight = .Width
        Cells.ColumnWidth = .ColumnWidth
    End With
End Sub


