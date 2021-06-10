Attribute VB_Name = "aMAIN"
Public S As Shape
Public cell As Range
Public rng As Range

'GAME OF BADUK (aka iGO or WeiQi)
'CODED BY ANASTASIOU ALEX
'ANASTASIOUALEX@GMAIL.COM

Sub SelectAdjacentSameValue()
    'https://www.mrexcel.com/board/threads/vba-macro-to-expand-selection-to-adjacent-cells-with-same-value.703756/
    Application.ScreenUpdating = False
    Dim Tbl As Range, Ctr As Range, rng As Range, Rtmp As Range, Ctest As Range, c As Range
    Dim v
    Dim Found As Boolean
'   Tbl = In which area to limit the search
    Set Tbl = Range("Goban")
'   Ctr = Where to start searching from (where we play the stone)
    Set Ctr = ActiveCell
'   If the start poin tis withing the search area then contitue
    If Not Intersect(Ctr, Tbl) Is Nothing Then
    
        Set Rtmp = Ctr
'       v = we search for cells with value v (we write B/W on placing a stone)
        v = Ctr.Value
        Do
            Set rng = Rtmp
            Found = False
            For Each c In rng
'            Search up, down, left, right to and add to the chain
'            the cells with the same value
                For Each Ctest In Union(c.Offset(-1), c.Offset(1), _
                                        c.Offset(, -1), c.Offset(, 1))
                    If Not Intersect(Ctest, Tbl) Is Nothing Then
                        If Intersect(Ctest, rng) Is Nothing Then
                            If Ctest.Value = v Then
                                Set Rtmp = Union(Rtmp, Ctest)
                                Found = True
                            End If
                        End If
                    End If
                Next Ctest
            Next c
        Loop While Found
        Application.EnableEvents = False
        rng.Select
        Ctr.Activate
        Application.EnableEvents = True
    End If
    Application.ScreenUpdating = True
End Sub

Sub GoSetup()
    'if the Mode = "Setup" then (this is checked within the GO sheet's code
    Application.ScreenUpdating = False
'   if not called from inside the board then exit sub
    If Intersect(ActiveCell, Range("Goban")) Is Nothing Then Application.ScreenUpdating = True: Exit Sub
'   If one cell is selected (The sub is called from worksheet selection change)
    If Selection.Cells.Count = 1 Then
'       place a stone for the active player
        ActiveCell = [Goturn]
'       duplicate the shape for the active player
        If [Goturn] = "W" Then
            Set S = ActiveSheet.Shapes("W").Duplicate
        Else: Set S = ActiveSheet.Shapes("B").Duplicate
        End If
'       move the duplicated shape to where we played
        With S
'           assign unique name to avoid errors when using application.caller
            .Name = .Name & " Duplicate " & ActiveSheet.Shapes.Count + 1
'           Display no text. these moves are not numbered.
            .TextFrame.Characters.Text = ""
'           call a sub to move, resize and center the shape
            FitShape
'           The original shape had a macro assigned. Remove it.
            .OnAction = ""
        End With
'       if more many cells selected
'       (it doesn't trigger ws selection change so it works regardless if mode is setup or game)
    
'       This part of the setup is to be used for creating puzzles.
'       It won't work with games because the stones here are not numbered moves.
    Else
'       For each cell in selection write W/B by clicking W/B stone next to MOVES
        Dim tmpcaller As String
        tmpcaller = ActiveSheet.Shapes(Application.Caller).Name
'       duplicate the correct shape for the active player
        For Each cell In Selection
'           Place stones only on the goban.
            If Not Intersect(cell, Range("GOBAN")) Is Nothing Then
                cell = tmpcaller
                If cell = "W" Then
                    Set S = ActiveSheet.Shapes("W").Duplicate
                Else: Set S = ActiveSheet.Shapes("B").Duplicate
                End If
                With S
'                   rename shape
                    .Name = .Name & " Duplicate " & ActiveSheet.Shapes.Count + 1
'                   Display no text. these moves are not numbered.
                    .TextFrame.Characters.Text = ""
'                   move, resize and center the shape
'                   we can't call the sub because all stones will be placed in the same cell
                    S.Height = cell.Height * 0.9
                    S.Width = cell.Width * 0.9
                    S.Left = cell.Left + ((cell.Width - S.Width) / 2)
                    S.Top = cell.Top + ((cell.Height - S.Height) / 2)
'                   The original shape had a macro assigned. Remove it.
                    .OnAction = ""
                End With
            End If
        Next cell
    End If
    If [GoLoop] = "" Then Application.ScreenUpdating = True
End Sub

Sub GoMove()
'   if the Mode = "Game" then (this is checked within the GO sheet's code
    Application.ScreenUpdating = False
    Dim cell As Range
    Dim MyChain As Range
    Dim Player As String
    Dim Enemy As String
    Dim Hommie As Range
'   remember where we played
'   the checking for enemy chains starts here
    Set Hommie = ActiveCell
'   place restriction of worksheet selection change
'   so we can interact with the board without calling actions
    [GoOperation] = "Skip"
'   begin the move by placing B or W for active player
    ActiveCell = [Goturn]
'   Detect whose turn it is
    Select Case [Goturn]
    Case "W"
'       define the enemy colour. We'll search for his chains.
        Enemy = "B"
'       duplicate the correct shape for the active player
        Set S = ActiveSheet.Shapes("W").Duplicate
'       Display inside the shape the turn when it was played
        S.TextFrame.Characters.Text = [CountMoveWhite] + [CountMoveBlack] + 3
'       Write down the move (where on the board it was placed)
'       so we can save a board position and replay the moves
        If [GoMovesWhite] = vbNullString Then
            [GoMovesWhite] = Hommie.Address
        Else: [GoMovesWhite] = [GoMovesWhite] & "," & Hommie.Address
        End If
'       Write down this move's number
        [CountMoveWhite] = [CountMoveWhite] + 1
    Case "B"
        Enemy = "W"
'       duplicate the correct shape for the active player
        Set S = ActiveSheet.Shapes("B").Duplicate
'       Display inside the shape the turn when it was played
        S.TextFrame.Characters.Text = [CountMoveBlack] + [CountMoveWhite] + 3
'       Write down the move (where on the board it was placed)
'       so we can save a board position and replay the moves
        If [GoMovesBlack] = vbNullString Then
            [GoMovesBlack] = Hommie.Address
        Else: [GoMovesBlack] = [GoMovesBlack] & "," & Hommie.Address
        End If
'       Write down this move's number
        [CountMoveBlack] = [CountMoveBlack] + 1
    End Select
'   move the duplicated shape to where we played
    With S
'       assign unique name to avoid errors when using application.caller
        .Name = .Name & " Duplicate " & ActiveSheet.Shapes.Count + 1
'       call a sub to move, resize and center the shape
        FitShape
'       The original shape had a macro assigned. Remove it.
        .OnAction = ""
    End With
'   check left, down, right, up from where we played
'   if there is an enemy stone then select the begining of the enemy chain
'   and call the sub GoCapture (see next sub)
    If Hommie.Offset(-1, 0) = Enemy Then
        Hommie.Offset(-1, 0).Select
        GoCapture
    End If
    If Hommie.Offset(1, 0) = Enemy Then
        Hommie.Offset(1, 0).Select
        GoCapture
    End If
    If Hommie.Offset(0, 1) = Enemy Then
        Hommie.Offset(0, 1).Select
        GoCapture
    End If
    If Hommie.Offset(0, -1) = Enemy Then
        Hommie.Offset(0, -1).Select
        GoCapture
    End If
'   switch active player (sub)
    GoSwitch
'   remove restriction of worksheet selection change
    [GoOperation] = vbNullString
    Application.ScreenUpdating = True
End Sub

Sub GoCapture()
    Application.ScreenUpdating = False
'   select the enemy chain
    SelectAdjacentSameValue
    Set MyChain = Selection
'   if any stone of the chain has a liberty (empty adjacent cell
'   then the chain is alive and not captured (exit sub)
    For Each cell In MyChain
        If Not Intersect(cell.Offset(-1, 0), Range("Goban")) Is Nothing And cell.Offset(-1, 0) = 0 Then Application.ScreenUpdating = True: Exit Sub
        If Not Intersect(cell.Offset(1, 0), Range("Goban")) Is Nothing And cell.Offset(1, 0) = 0 Then Application.ScreenUpdating = True: Exit Sub
        If Not Intersect(cell.Offset(0, 1), Range("Goban")) Is Nothing And cell.Offset(0, 1) = 0 Then Application.ScreenUpdating = True: Exit Sub
        If Not Intersect(cell.Offset(0, -1), Range("Goban")) Is Nothing And cell.Offset(0, -1) = 0 Then Application.ScreenUpdating = True: Exit Sub
    Next
'   if there are no liberties then remove the chain
'   When the board is clear, the cells' value is 0.
'   when they are W/B they are counted at the EndGame for scoring
'   ClearContents
    MyChain.Value = 0
'   remove all enemy stones from the board
    For Each S In ActiveSheet.Shapes
        If Not Intersect(Range(S.TopLeftCell.Address), MyChain) Is Nothing Then S.Delete
    Next S
'   count the captured stones.
'   This is used in a scoring method different from what I use. The result is the same.
'   Since we're here let's have this feature too.
    Select Case [Goturn]
    Case "B"
        [CapturedBlack] = [CapturedBlack] + MyChain.Cells.Count
    Case "W"
        [CapturedWhite] = [CapturedWhite] + MyChain.Cells.Count
    End Select
    Application.ScreenUpdating = True
End Sub

Sub GoSwitch()
'   change the active player
'   assigned to the single stone next to TURN
    If [Goturn] = "W" Then
        [Goturn] = "B"
        ActiveSheet.Shapes("GoWhiteTurn").Visible = False
        ActiveSheet.Shapes("GoBlackTurn").Visible = True
    Else: [Goturn] = "W"
        ActiveSheet.Shapes("GoBlackTurn").Visible = False
        ActiveSheet.Shapes("GoWhiteTurn").Visible = True
    End If
End Sub


