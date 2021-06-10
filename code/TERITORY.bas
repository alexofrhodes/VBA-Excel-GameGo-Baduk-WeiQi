Attribute VB_Name = "TERITORY"
Sub EndGame()
    If [GoOperation] = vbNullString Then
        [GoOperation] = "EndGame"
        MsgBox "Switch active player as necessary." & _
               Chr(10) & "Remove dead stones (click on them) and" & _
               Chr(10) & "attribute territory (click on empty SURROUNDED space."
'       assign sub to shapes on goban
        For Each S In ActiveSheet.Shapes
            If Not Intersect(Range(S.TopLeftCell.Address), Range("Goban")) Is Nothing Then
                S.OnAction = "MarkTerritory"
            End If
        Next S
        Cells(1, ActiveCell.Column).Select
    ElseIf [GoOperation] = "EndGame" Then
        [ScoreWhite] = [ScoreWhite] + [komi]
        If [ScoreBlack] > [ScoreWhite] Then MsgBox "Winner is Black by " & [ScoreBlack] - [ScoreWhite] & " points."
        If [ScoreBlack] < [ScoreWhite] Then MsgBox "Winner is White by " & [ScoreWhite] - [ScoreBlack] & " points."
        If [ScoreBlack] = [ScoreWhite] Then MsgBox "It is a draw."
        Cells(1, ActiveCell.Column).Select
        [GoOperation] = vbNullString
    End If
End Sub

Sub AttributeTerritory()
    Application.ScreenUpdating = False
    [GoMode] = "Setup"
    SelectAdjacentSameValue
    Dim TEMP As Range
    Set TEMP = ActiveCell
    SelectAdjacentSameValue
    Set MyChain = Selection
    MyChain.Value = [Goturn]
    [GoOperation] = ""
    [A1].Select
    For Each cell In MyChain
        cell.Select
    Next cell
    Select Case [Goturn]
    Case "B"
        [ScoreBlack] = [ScoreBlack] + MyChain.Cells.Count
    Case "W"
        [ScoreWhite] = [ScoreWhite] + MyChain.Cells.Count
    End Select
    [GoOperation] = "EndGame"
    Application.ScreenUpdating = True
End Sub

Sub MarkTerritory()
    Application.ScreenUpdating = False
'   MarkDead
    [GoMode] = "Setup"
    [GoOperation] = "Skip"
    Dim TEMP As Range
    Set TEMP = ActiveSheet.Shapes(Application.Caller).TopLeftCell
    TEMP.Select
    SelectAdjacentSameValue
    Set MyChain = Selection
    For Each S In ActiveSheet.Shapes
        If Not Intersect(Range(S.TopLeftCell.Address), MyChain) Is Nothing Then S.Delete '  s.Select Replace:=False
    Next S
    MyChain.Value = 0
    TEMP.Select
    SelectAdjacentSameValue
    Set MyChain = Selection
    MyChain.Value = [Goturn]
    [A1].Select
    [GoOperation] = ""
    For Each cell In MyChain
        cell.Select
    Next cell
    Select Case [Goturn]
    Case "B"
        [ScoreBlack] = [ScoreBlack] + MyChain.Cells.Count
    Case "W"
        [ScoreWhite] = [ScoreWhite] + MyChain.Cells.Count
    End Select
    [GoOperation] = "EndGame"
    Application.ScreenUpdating = True
End Sub
