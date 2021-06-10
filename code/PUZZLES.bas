Attribute VB_Name = "PUZZLES"
Sub PuzzleInit()
    Application.ScreenUpdating = True
    If [TotalPuzzles] = 0 Then MsgBox "No puzzle found.": Exit Sub
    Application.ScreenUpdating = False
    Dim PuzzleData As Range
'   filter puzzles in place
    Sheets("PUZZLES").Range("A4").CurrentRegion.AdvancedFilter _
        Action:=xlFilterInPlace, _
        CriteriaRange:=Sheets("GO").Range("CriteriaPuzzle")
    If IsNumeric(Sheets("PUZZLES").Range("A" & Rows.Count).End(xlUp)) = False Then Application.ScreenUpdating = True: MsgBox "No Puzzles stored.": Exit Sub
'   get random from filtered
    Dim tmp As Range
    Set tmp = Sheets("PUZZLES").Range(RndVariable("PUZZLES"))
    Range(tmp, tmp.Offset(0, [FilteredPuzzle].Columns.Count - 1)).Copy _
        Sheets("GO").Range("FilteredPuzzle").Offset(1, 0)
    Sheets("GO").Range("FilteredPuzzle").Offset(1, 0).Resize(1).HorizontalAlignment = xlCenter
    Sheets("GO").Range("FilteredPuzzle").Offset(1, 0).Resize(1).VerticalAlignment = xlCenter
'   clear criteria id but keep other manualy set criteria
    Range("CriteriaPuzzle").Offset(1, 0).Resize(1, 1).ClearContents
    PuzzleReload
    Application.ScreenUpdating = True
End Sub

Sub PuzzleReload()
    [ksize] = [pKsize]
    If [ksize] = 9 And Columns(11).Hidden = False Then GobanResize
    If [ksize] = 13 And Columns(15).Hidden Or Columns(20).Hidden = False Then GobanResize
    If [ksize] = 19 And Columns(20).Hidden Then GobanResize
    Application.ScreenUpdating = False
    GoReset
    Sheets("GO").Range("GoMovesBlack") = Range("FilteredPuzzle").Offset(1, 1).Resize(1, 1)
    Sheets("GO").Range("GoMovesWhite") = Range("FilteredPuzzle").Offset(1, 2).Resize(1, 1)
    PuzzleLoad
    Application.ScreenUpdating = True
End Sub

Sub PuzzleLoad()
    Application.ScreenUpdating = False
    Dim arr As Variant
    Dim i As Long
'    On Error Resume Next
    [GoMode] = "Setup"
    If [Goturn] = "W" Then GoSwitch
    arr = Split([GoMovesBlack], ",")
    For i = LBound(arr) To UBound(arr)
        Range(arr(i)).Select
    Next i
'    [CountMoveBlack] = UBound(arr)
    [CountMoveBlack] = -1
    [GoMovesBlack] = ""
    GoSwitch
    arr = Split([GoMovesWhite], ",")
    For i = LBound(arr) To UBound(arr)
        Range(arr(i)).Select
    Next i
'    [CountMoveWhite] = UBound(arr)
    [CountMoveWhite] = -1
    [GoMovesWhite] = ""
    [FilteredPuzzle].Offset(1, 0).Select
    PasteTextFormat
    Range("$AB$5:$AH$5,$AB$8:$AH$8,$AB$14:$AK$14,$AB$17:$AK$17,$Z$11:$AA$11,$Z$2:$AA$2,$W$1:$X$1,$X$4,$W$8:$X$9,$W$13:$X$13,$W$15:$X$15,$W$17").Select
    JustifyRange
    Cells(1, ActiveCell.Column).Select
    [gLoaded] = ""
    [pLoaded] = "LOADED"
    If [Goturn] = "W" Then GoSwitch
    [GoMode] = "Game"
    Application.ScreenUpdating = True
End Sub

Sub PuzzleAddNew()
    Application.ScreenUpdating = False
'   Add new puzzle from board position
    [GoMovesBlack].ClearContents
    [GoMovesWhite].ClearContents
    [GoMode] = "Setup"
    For Each cell In Range("Goban")
        If cell = "B" Then
            If [GoMovesBlack] = vbNullString Then
                [GoMovesBlack] = cell.Address
            Else: [GoMovesBlack] = [GoMovesBlack] & "," & cell.Address
            End If
        End If
        If cell = "W" Then
            If [GoMovesWhite] = vbNullString Then
                [GoMovesWhite] = cell.Address
            Else: [GoMovesWhite] = [GoMovesWhite] & "," & cell.Address
            End If
        End If
    Next cell
    Range("FilteredPuzzle").Offset(1, 0).ClearContents
    [pKsize] = [ksize]
    [FilteredPuzzle].Offset(1, 0).Resize(1, 1) = [TotalPuzzles] + 1
    [FilteredPuzzle].Offset(1, 1).Resize(1, 1) = [GoMovesBlack]
    [FilteredPuzzle].Offset(1, 2).Resize(1, 1) = [GoMovesWhite]
    Range("$AB$5:$AH$5,$AB$8:$AH$8,$AB$14:$AK$14,$AB$17:$AK$17,$Z$11:$AA$11,$Z$2:$AA$2,$W$1:$X$1,$X$4,$W$8:$X$9,$W$13:$X$13,$W$15:$X$15,$W$17").Select
    JustifyRange
    PuzzleUpdate
    PuzzleReload
    Application.ScreenUpdating = True
End Sub

Sub PuzzleUpdate()
    Application.ScreenUpdating = False
    Select Case [FilteredPuzzle].Offset(1, 0).Resize(1, 1)
    Case Is <= [TotalPuzzles]
        [FilteredPuzzle].Offset(1, 0).Resize(1).Copy
        Sheets("PUZZLES").Range("A:A").Find( _
        [FilteredPuzzle].Offset(1, 0).Resize(1, 1)).PasteSpecial Paste:=xlPasteValues
    Case Is > [TotalPuzzles]
        If IsNumeric(Sheets("PUZZLES").Range("A" & Rows.Count).End(xlUp)) = False Then
            Sheets("PUZZLES").Range("A" & Rows.Count).End(xlUp).Offset(1, 0) = 1
        Else
            Sheets("PUZZLES").Range("A" & Rows.Count).End(xlUp).Offset(1, 0) = [TotalPuzzles] + 1
        End If
        [FilteredPuzzle].Offset(1, 0).Resize(1).Copy
        Sheets("PUZZLES").Range("A" & Rows.Count).End(xlUp).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End Select
    Cells(1, ActiveCell.Column).Select
    Application.CutCopyMode = False
    Cells(1, ActiveCell.Column).Select
    Application.ScreenUpdating = True
    MsgBox "Puzzle saved / updated."
End Sub

Sub DeletePuzzle()
    Application.ScreenUpdating = True
    Cells(1, ActiveCell.Column).Select
    If [FilteredPuzzle].Offset(1, 0).Resize(1, 1) = "" Then Application.ScreenUpdating = True: MsgBox "Nothing to delete.": Exit Sub
    Dim i As Long
    Dim confirmation As Integer
    confirmation = MsgBox("Are you sure you want to delete this Puzzle?", vbQuestion + vbYesNo + vbDefaultButton1, "File saved")
    If confirmation = vbNo Then Application.ScreenUpdating = True: Exit Sub
    Application.ScreenUpdating = False
    i = 0
    Sheets("PUZZLES").Range("A:A").Find( _
        [FilteredPuzzle].Offset(1, 0).Resize(1, 1)).EntireRow.Delete
    If Sheets("PUZZLES").Range("A5") <> "" Then
        For Each cell In Sheets("PUZZLES").Range("A5:" & Sheets("PUZZLES").Range("A" & Rows.Count).End(xlUp).Address)
            i = i + 1
            cell = i
        Next cell
    End If
    Sheets("GO").Range("FilteredPuzzle").Offset(1, 0).ClearContents
    Application.ScreenUpdating = True
    MsgBox "Puzzle deleted."
End Sub

Sub PuzzleClear()
    [FilteredPuzzle].Offset(1, 0).ClearContents
End Sub
