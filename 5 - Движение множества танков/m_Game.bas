Attribute VB_Name = "m_Game"
Public isStop As Boolean
Private tanks As Collection











Public Sub Game()

Dim shp As Visio.Shape
Dim speed As Double
Dim x As Double
Dim y As Double
Dim direction As Double
Dim endTime As Date

    isStop = False
    endTime = DateAdd("s", 60, Now())
    
    Set tanks = New Collection
    FindTanks
    

    Do While endTime > Now()
        For Each shp In tanks
            speed = shp.Cells("Prop.Speed").Result(visNumber)
            direction = shp.Cells("Angle").Result(visRadians) + rad90
            
            
            x = shp.Cells("PinX").Result(visInches) + speed * Cos(direction)
            y = shp.Cells("PinY").Result(visInches) + speed * Sin(direction)
            shp.Cells("PinX").Formula = x
            shp.Cells("PinY").Formula = y
        Next shp

        DoEvents

        If isStop = True Then
            Exit Sub
        End If
        
    Loop
    


End Sub


Private Sub FindTanks()
Dim shp As Visio.Shape
    
    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.GameObject", 0) <> 0 Then
            If shp.Cells("User.GameObject").Result(visNumber) = 1 Then
                tanks.Add shp
            End If
        End If
    Next shp
    
End Sub
