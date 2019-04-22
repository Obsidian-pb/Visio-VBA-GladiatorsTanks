VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Shell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public id As String
Public shp As Visio.Shape
Public currentPosition As c_Point
Const speed As Double = 0.1
Public direction As Double
Public tanks As Collection
Public firedTank As c_Tank
Private i As Long



Public Function Render() As Boolean
    currentPosition.x = currentPosition.x + speed * Cos(direction)
    currentPosition.y = currentPosition.y + speed * Sin(direction)
    
    shp.Cells("PinX").Formula = currentPosition.x
    shp.Cells("PinY").Formula = currentPosition.y
    
    If currentPosition.x < 0 Or currentPosition.x > Application.ActivePage.PageSheet.Cells("PageWidth").Result(visInches) Or _
        currentPosition.y < 0 Or currentPosition.y > Application.ActivePage.PageSheet.Cells("PageHeight").Result(visInches) Then
        Render = True
    End If
    
    i = i + 1       'Preserve self demolition
    If i < 10 Then Exit Function
    
    On Error GoTo EX
    Dim tank As c_Tank
    For Each tank In tanks
        If tank.shp.HitTest(currentPosition.x, currentPosition.y, 0) <> 0 Then
            tanks.Remove tank.id
            tank.shp.Delete
            TargetsClear tank.id
            Set tank = Nothing
            
            firedTank.RizeScore
            
            Render = True
            Exit Function
        End If
    Next tank
    
EX:
    
    
End Function


Public Sub SetLayer(ByVal LayerName As String)
    shp.CellsSRC(visSectionObject, visRowLayerMem, visLayerMember).FormulaU = GetLayerNumber(LayerName)
End Sub

Private Sub TargetsClear(ByRef targetID As String)
Dim tank As c_Tank

    For Each tank In tanks
        If tank.target.id = targetID Then Set tank.target = Nothing
    Next tank

End Sub
