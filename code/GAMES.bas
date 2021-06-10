Attribute VB_Name = "GAMES"
Global tmpcap As String

Sub GameInit()
'   exit sub if no games saved
    If [TotalGames] = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No game found."
        Exit Sub
    End If
    Dim GameData As Range
'   filter GAMES in place
    Sheets("GAMES").Range("A4").CurrentRegion.AdvancedFilter _
        Action:=xlFilterInPlace, _
        CriteriaRange:=Sheets("GO").Range("CriteriaGame")
'    if no game matches the criteria exit sub
    If IsNumeric(Sheets("GAMES").Range("A" & Rows.Count).End(xlUp)) = False Then
        MsgBox "No GAMES found."
        Exit Sub
    End If
    Application.ScreenUpdating = False
'   get random from filtered
    Dim tmp As Range
    Set tmp = Sheets("GAMES").Range(RndVariable("GAMES"))
    Range(tmp, tmp.Offset(0, [FilteredGame].Columns.Count - 1)).Copy _
        Sheets("GO").Range("FilteredGame").Offset(1, 0)
'    Import loaded game's handicap
    Sheets("GO").Range("FilteredGame").Offset(1, 0).Resize(1).HorizontalAlignment = xlCenter
    Sheets("GO").Range("FilteredGame").Offset(1, 0).Resize(1).VerticalAlignment = xlCenter
'   clear criteria id but keep other manualy set criteria
    Range("CriteriaGame").Offset(1, 0).Resize(1, 1).ClearContents
'   add the data on the board
    GameReload
    If Application.ScreenUpdating = False Then: Application.ScreenUpdating = True
End Sub

Sub GameReload()
    If Application.ScreenUpdating = True Then: Application.ScreenUpdating = False
'    Load game's board size and resize if needed
    [ksize] = [gKsize]
    If [ksize] = 9 And Columns(11).Hidden = False Then GobanResize
    If [ksize] = 13 And Columns(15).Hidden Or Columns(20).Hidden = False Then GobanResize
    If [ksize] = 19 And Columns(20).Hidden Then GobanResize
'   Check if the game being loaded was handicapped
    If [gsetup] = "" Then
        GoReset
    Else
        handicap = [gsetup]
        [WHATCAP] = [gsetup]
        HandicappedGame
    End If
'    Import moves
    Sheets("GO").Range("GoMovesBlack") = Range("FilteredGame").Offset(1, 1).Resize(1, 1)
    Sheets("GO").Range("GoMovesWhite") = Range("FilteredGame").Offset(1, 2).Resize(1, 1)
    GameLoad
    Application.ScreenUpdating = True
End Sub

Sub GameLoad()
    Dim arr1    As Variant
    Dim arr2    As Variant
    Dim arr3    As Variant
    Dim cRng    As String
    Dim i       As Long
    Dim x       As Long
    Dim a       As Long
    Dim B       As Long
    Dim tmp1    As String
    Dim tmp2    As String
    
    Cells(1, ActiveCell.Column).Select
    [GoMode] = "Game"
    cRng = ""
    a = [CountMoveBlack]
    B = [CountMoveWhite]
    tmp1 = [GoMovesBlack]
    tmp2 = [GoMovesWhite]
'   Create 2 arrays for B and W moves
    arr1 = Split([GoMovesBlack], ",")
    arr2 = Split([GoMovesWhite], ",")
'   check if the game being loaded is handicapped
'   to find which player moved first and create the mixed array cRng.
    Select Case [komi]
'    if it is not handicapped then black moved first
    Case Is > 0.5
HANDICAP_1:
'       For i = 0 to [CountMoveBlack]
        For i = LBound(arr1) To UBound(arr1)
'           Add BlackMove i to the temp string
            cRng = cRng & "," & arr1(i)
'           If i <= Last WhiteMove then add WhiteMove i to the string
'           if i > then there would be no WhiteMove to add
            If i <= UBound(arr2) Then cRng = cRng & "," & arr2(i)
        Next i
'    if it is handicapped then white moved first
    Case Is = 0.5
'        exception: if handicap = 1 then black moved first
        If handicap = "1" Then GoTo HANDICAP_1
'       For i = 0 to [CountMoveWhite]
        For i = LBound(arr2) To UBound(arr2)
'           Add WhiteMove i to the temp string
            cRng = cRng & "," & arr2(i)
'           If i <= Last BlackMove then add BlackMove i to the string
'           if i > then there would be no BlackMove to add
            If i <= UBound(arr1) Then cRng = cRng & "," & arr1(i)
        Next i
    End Select
    
    On Error Resume Next
'   remove the , from the start of the string
    cRng = Right(cRng, Len(cRng) - 1)
'   now we have an alternating mix of B and W moves
    arr3 = Split(cRng, ",")
'   Total number of moves
    x = a + B
'   reset the board and moves count
    Range("Goban").Value = 0
    For Each S In ActiveSheet.Shapes
        Select Case [WHATCAP]
        Case Is <> ""
            If Not Intersect(Range(S.TopLeftCell.Address), Range("Goban")) Is Nothing _
            And S.TextFrame.Characters.Text <> "W" And S.TextFrame.Characters.Text <> "B" _
            And S.TextFrame.Characters.Text <> "" Then
                S.Delete
            End If
        Case Else
            If Not Intersect(Range(S.TopLeftCell.Address), Range("Goban")) Is Nothing Then
                S.Delete
            End If
        End Select
    Next S
    [CountMoveWhite] = -1
    [CountMoveBlack] = -1
'    make the correct player active
    If [komi] > 0.5 Or handicap = "1" Then
'       make sure we start with B
        If [Goturn] = "W" Then GoSwitch
    ElseIf [komi] = 0.5 Then
'       make sure we start with W
        If [Goturn] = "B" Then GoSwitch
    End If
'    play out the moves
'   just setting up the board based on the saved moves
'   wouldn't work because of capture moves.
    i = 0
    For i = LBound(arr3) To UBound(arr3)         'x + 1
        Range(arr3(i)).Select
    Next i
    Cells(1, ActiveCell.Column).Select
'   Write back the original moves so we can continue to undo/redo them
    [GoMovesBlack] = tmp1
    [GoMovesWhite] = tmp2
    cRng = ""
    [FilteredGame].Offset(1, 0).Select
    PasteTextFormat
    Range("$AB$5:$AH$5,$AB$8:$AH$8,$AB$14:$AK$14,$AB$17:$AK$17,$Z$11:$AA$11,$Z$2:$AA$2,$W$1:$X$1,$X$4,$W$8:$X$9,$W$13:$X$13,$W$15:$X$15,$W$17").Select
    JustifyRange
    Cells(1, ActiveCell.Column).Select
    [gLoaded] = "LOADED"
    [pLoaded] = ""
End Sub

Sub GameAddNew()
    Application.ScreenUpdating = False
'    clear the range to add the new game data
    Range("FilteredGame").Offset(1, 0).ClearContents
'    save board size
    [gKsize] = [ksize]
'    handicap if any
    [gsetup] = [WHATCAP]
'    assign game number
    [FilteredGame].Offset(1, 0).Resize(1, 1) = [TotalGames] + 1
'    import B and W moves
    [FilteredGame].Offset(1, 1).Resize(1, 1) = [GoMovesBlack]
    [FilteredGame].Offset(1, 2).Resize(1, 1) = [GoMovesWhite]
    Range("$AB$5:$AH$5,$AB$8:$AH$8,$AB$14:$AK$14,$AB$17:$AK$17,$Z$11:$AA$11,$Z$2:$AA$2,$W$1:$X$1,$X$4,$W$8:$X$9,$W$13:$X$13,$W$15:$X$15,$W$17").Select
    JustifyRange
'    Save the new game
    GameUpdate
    GameReload
End Sub

Sub GameUpdate()
    If Application.ScreenUpdating = True Then: Application.ScreenUpdating = False
'    check if adding a new game or updating an existing one
    Select Case [FilteredGame].Offset(1, 0).Resize(1, 1)
'    updating game
    Case Is <= [TotalGames]
'        overwrite old data with new
        [FilteredGame].Offset(1, 0).Resize(1).Copy
        Sheets("GAMES").Range("A:A").Find( _
        [FilteredGame].Offset(1, 0).Resize(1, 1)).PasteSpecial Paste:=xlPasteValues
'    adding new game
    Case Is > [TotalGames]
'    add game's number
'        if first game
        If IsNumeric(Sheets("GAMES").Range("A" & Rows.Count).End(xlUp)) = False Then
            Sheets("GAMES").Range("A" & Rows.Count).End(xlUp).Offset(1, 0) = 1
        Else
            Sheets("GAMES").Range("A" & Rows.Count).End(xlUp).Offset(1, 0) = [TotalGames] + 1
        End If
'        add new game's data
        [FilteredGame].Offset(1, 0).Resize(1).Copy
        Sheets("GAMES").Range("A" & Rows.Count).End(xlUp).PasteSpecial Paste:=xlPasteValues
    End Select
    Application.CutCopyMode = False
    Cells(1, ActiveCell.Column).Select
    Application.ScreenUpdating = True
    MsgBox "Game saved / updated."
End Sub

Sub DeleteGame()
    Cells(1, ActiveCell.Column).Select
    If [FilteredGame].Offset(1, 0).Resize(1, 1) = "" Then
        MsgBox "No game loaded to delete it."
        Exit Sub
    End If
    Dim i As Long
    Dim confirmation As Integer
'    confirm delete
    confirmation = MsgBox("Are you sure you want to delete this game?", _
        vbQuestion + vbYesNo + vbDefaultButton1)
    If confirmation = vbNo Then Exit Sub
    Application.ScreenUpdating = False
'    Detete the game row
    Sheets("GAMES").Range("A:A").Find( _
        [FilteredGame].Offset(1, 0).Resize(1, 1)).EntireRow.Delete
'    Renumber the games
    i = 0
    If Sheets("GAMES").Range("A5") <> "" Then
        For Each cell In Sheets("GAMES").Range("A5:" & Sheets("GAMES").Range("A" & Rows.Count).End(xlUp).Address)
            i = i + 1
            cell = i
        Next cell
    End If
    Sheets("GO").Range("FilteredGame").Offset(1, 0).ClearContents
    Application.ScreenUpdating = True
'    MsgBox "Game deleted."
End Sub

Sub GameClear()
    [FilteredGame].Offset(1, 0).ClearContents
End Sub


