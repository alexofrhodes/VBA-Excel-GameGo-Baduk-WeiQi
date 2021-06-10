Attribute VB_Name = "CONTROLS"
Sub Lcontrols()
    If [pLoaded] <> "" Then
    Dim answer As Integer
    answer = MsgBox("Replay is intended for reviewing Games." & _
    vbNewLine & "Do you want to reload this Puzzle instead?", vbQuestion + vbYesNo + vbDefaultButton1, "Message Box Title")
        If answer = vbYes Then
            Application.ScreenUpdating = False
            PuzzleReload
            Exit Sub
        Else
        Exit Sub
        End If
    End If
'   Switch to Game mode. We'll need to make moves.
    [GoMode] = "Game"
'   If we're at the beginning with no more moves, msg and exit sub
    If [CountMoveBlack] = -1 And [CountMoveWhite] = -1 Then MsgBox "No more moves.": Cells(1, ActiveCell.Column).Select: Exit Sub
    Application.ScreenUpdating = False
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
    handicap = [WHATCAP]
    Cells(1, ActiveCell.Column).Select
    cRng = ""
    a = [CountMoveBlack]
    B = [CountMoveWhite]
    tmp1 = [GoMovesBlack]
    tmp2 = [GoMovesWhite]
'   Create 2 arrays for B and W moves
    arr1 = Split([GoMovesBlack], ",")
    arr2 = Split([GoMovesWhite], ",")
    Select Case [komi]
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
    Case Is = 0.5
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
'   remove the , from the start of the string
    cRng = Right(cRng, Len(cRng) - 1)
'   now we have an alternating mix of B and W moves
    arr3 = Split(cRng, ",")
'   Total number of moves
    x = a + B
'   reset the board and moves count
    Range("Goban").Value = 0
    For Each S In ActiveSheet.Shapes
        If Not Intersect(Range(S.TopLeftCell.Address), Range("Goban")) Is Nothing _
            And S.TextFrame.Characters.Text <> "W" And S.TextFrame.Characters.Text <> "B" _
            And S.TextFrame.Characters.Text <> "" Then S.Delete '  s.Select Replace:=False
    Next S
    [CountMoveWhite] = -1
    [CountMoveBlack] = -1
    
    Select Case [komi]
    Case Is > 0.5
HANDICAP_2:
'       make sure we start with B
        If [Goturn] = "W" Then GoSwitch
    Case Is = 0.5
    If handicap = "1" Then GoTo HANDICAP_2
'       make sure we start with W
        If [Goturn] = "B" Then GoSwitch
    End Select
    i = 0
    For i = LBound(arr3) To x
        Range(arr3(i)).Select
    Next i
    Cells(1, ActiveCell.Column).Select
'   Write back the original moves so we can continue to undo/redo them
'   If you want to start a variation from the changed position then click NEW to save a new game or puzzle.
    [GoMovesBlack] = tmp1
    [GoMovesWhite] = tmp2
    cRng = ""
    Application.ScreenUpdating = True
End Sub

Sub Rcontrols()
    If [pLoaded] <> "" Then
    Dim answer As Integer
    answer = MsgBox("Replay is intended for reviewing Games." & _
    vbNewLine & "Do you want to reload this Puzzle instead?", vbQuestion + vbYesNo + vbDefaultButton1, "Message Box Title")
        If answer = vbYes Then
            Application.ScreenUpdating = False
            PuzzleReload
            Exit Sub
        Else
        Exit Sub
        End If
    End If
'   Switch to Game mode. We'll need to make moves.
    [GoMode] = "Game"
    handicap = [WHATCAP]
'   Error if trying to play a move that doesn't exist. msg and exit sub
    On Error GoTo eh
    Application.ScreenUpdating = False
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
    tmp1 = [GoMovesBlack]
    tmp2 = [GoMovesWhite]
    a = [CountMoveBlack]
    B = [CountMoveWhite]
    cRng = ""
'   Total number of moves (we start from -1 and -1 for the undo done before)
    x = a + B + 2
    arr1 = Split([GoMovesBlack], ",")
    arr2 = Split([GoMovesWhite], ",")
    Select Case [komi]
    Case Is > 0.5
HANDICAP_2:
'       For i = 0 to [CountMoveBlack]
        For i = LBound(arr1) To UBound(arr1)
'           Add BlackMove i to the temp string
            cRng = cRng & "," & arr1(i)
'           If i <= Last WhiteMove then add WhiteMove i to the string
'           if i > then there would be no WhiteMove to add
            If i <= UBound(arr2) Then cRng = cRng & "," & arr2(i)
        Next i
    Case Is = 0.5
    If handicap = "1" Then GoTo HANDICAP_2
'       For i = 0 to [CountMoveWhite]
        For i = LBound(arr2) To UBound(arr2)
'           Add WhiteMove i to the temp string
            cRng = cRng & "," & arr2(i)
'           If i <= Last BlackMove then add BlackMove i to the string
'           if i > then there would be no BlackMove to add
            If i <= UBound(arr1) Then cRng = cRng & "," & arr1(i)
        Next i
    End Select
'   remove the , from the start of the string
    cRng = Right(cRng, Len(cRng) - 1)
'   now we have an alternating mix of B and W moves
    arr3 = Split(cRng, ",")
'   solved another way?
    If x > UBound(arr3) Then GoTo eh
'   make the move
    Range(arr3(x)).Select
'   Write back the original moves so we can continue to undo/redo them
    [GoMovesBlack] = tmp1
    [GoMovesWhite] = tmp2
    cRng = ""
    Exit Sub
eh:
    Application.ScreenUpdating = True
    MsgBox "No more moves."
End Sub


