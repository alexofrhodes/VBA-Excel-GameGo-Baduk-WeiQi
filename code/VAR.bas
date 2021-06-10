Attribute VB_Name = "VAR"
Function RndVariable(sht As String)
'called from subs
    Dim tmprng As Variant
    Set tmprng = Sheets(sht).Range("A5:A999999").SpecialCells(xlCellTypeConstants)
    Set tmprng = tmprng.SpecialCells(xlCellTypeVisible)
    Dim tmpstr As String
    For Each cell In tmprng
        tmpstring = tmpstring & "," & cell.Address
    Next
    tmpstring = Right(tmpstring, Len(tmpstring) - 1)
    Dim tmpvar As Variant
    tmpvar = Split(tmpstring, ",")
    RndVariable = tmpvar(Int((UBound(tmpvar) + 1) * Rnd))
    Debug.Print RndVariable
End Function

Sub DeleteFromSheet()
    'CODE CALLED IN ANOTHER SUB
    If Selection.Row < 5 Then MsgBox "Please select cell(s) after row 4": Exit Sub
    Dim i As Long
    Dim confirmation As Integer
    confirmation = MsgBox("Are you sure you want to delete this?", vbQuestion + vbYesNo + vbDefaultButton1, "File saved")
    If confirmation = vbNo Then Exit Sub
    Application.ScreenUpdating = False
    Dim sel As Range
    Set sel = Selection.SpecialCells(xlVisible)
    sel.EntireRow.Delete
    If ActiveSheet.Name = "PUZZLES" Then
        Sheets("GO").Range("FilteredPuzzle").Offset(1, 0).ClearContents
    ElseIf ActiveSheet.Name = "GAMES" Then
        Sheets("GO").Range("FilteredGame").Offset(1, 0).ClearContents
    End If
    If WorksheetFunction.CountA(Range("A5:" & Sheets(ActiveSheet.Name).Range("A" & Rows.Count).End(xlUp).Address)) = 0 Then Exit Sub
    Rows.EntireRow.Hidden = False
    If [A5] = "" Then Exit Sub
    i = 0
    For Each cell In Range("A5:" & Sheets(ActiveSheet.Name).Range("A" & Rows.Count).End(xlUp).Address)
        i = i + 1
        cell = i
    Next cell
    Application.ScreenUpdating = True
End Sub

Sub ShapesColorByApplicationCaller()
    'CODE CALLED from shape
    For Each S In ActiveSheet.Shapes
        Select Case S.Name
        Case Is = "B", "W", "GoBlackTurn", "ShapeSample", "GoWhiteTurn"
            'do nothing
        Case Else
            ActiveSheet.Shapes("ShapeSample").PickUp
            S.Apply
        End Select
    Next
End Sub

Sub FitShape()
    'CODE CALLED IN ANOTHER SUB
    S.Height = ActiveCell.Height * 0.9
    S.Width = ActiveCell.Width * 0.9
    S.Left = ActiveCell.Left + ((ActiveCell.Width - S.Width) / 2)
    S.Top = ActiveCell.Top + ((ActiveCell.Height - S.Height) / 2)
End Sub

Sub JustifyRange()
    'CODE CALLED IN ANOTHER SUB
    Application.ScreenUpdating = False
    With Selection
        .HorizontalAlignment = xlJustify
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Application.ScreenUpdating = True
End Sub

Sub ShapesFormat()
    For Each S In ActiveSheet.Shapes
        'For Each S In Selection.ShapeRange
        'S.Name = S.TextFrame.Characters.Text
        S.Height = S.TopLeftCell.Height * 0.9
        S.Width = S.TopLeftCell.Width * 0.9
        S.Left = S.TopLeftCell.Left + ((S.TopLeftCell.Width - S.Width) / 2)
        S.Top = S.TopLeftCell.Top + ((S.TopLeftCell.Height - S.Height) / 2)
    Next S
End Sub

Sub PasteTextFormat()
    'CODE CALLED IN ANOTHER SUB
    Application.ScreenUpdating = False
    [fText].Copy
    Selection.PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Cells(1, ActiveCell.Column).Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

Sub CopyFormat()
    'CODE CALLED IN ANOTHER SUB
    Application.ScreenUpdating = False
    Range(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Address).Offset(0, -2).Copy
    Range(Range(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Address).Offset(0, -1).Value).Select
    Selection.PasteSpecial _
        Paste:=xlPasteFormats, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Cells(1, ActiveCell.Column).Select
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

Sub StarPoints()
    Application.ScreenUpdating = False
    Range("$E$5,$K$5,$Q$5,$Q$11,$Q$17,$K$17,$K$11,$E$11,$E$17").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Application.ScreenUpdating = True
End Sub

Sub ButtonMaker()
    Application.ScreenUpdating = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 22
        .Bold = True
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -9.99481185338908E-02
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -9.99481185338908E-02
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThick
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249946592608417
        .PatternTintAndShade = 0
    End With
    If Selection.Cells.Count > 1 Then Call ButtonsMaker
    Application.ScreenUpdating = True
End Sub

Sub ButtonsMaker()
    Application.ScreenUpdating = False
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 3
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Application.ScreenUpdating = True
End Sub

Sub WriteRangeAddress()
    'CODE CALLED FROM SHAPE
    [SelectionAddress] = Selection.Address
End Sub

Sub CodeSwitch()
    With ActiveSheet.Shapes(Application.Caller).TopLeftCell.Offset(0, 1).Resize(1, 2).EntireColumn
        If .Hidden = True Then
            .Hidden = False
        Else: .Hidden = True
        End If
    End With
End Sub

Sub ClearGPfilters()
    'CALLED FROM SHAPE
    [A1].CurrentRegion.Offset(1, 0).ClearContents
End Sub


