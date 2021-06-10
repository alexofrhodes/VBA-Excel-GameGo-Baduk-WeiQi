VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo eh
    If Not Intersect(Target, [A2:F2]) Is Nothing Then
        [A4].CurrentRegion.AdvancedFilter _
        Action:=xlFilterInPlace, _
        CriteriaRange:=Cells(1, ActiveCell.Column).CurrentRegion
    End If
eh:         Exit Sub
End Sub


