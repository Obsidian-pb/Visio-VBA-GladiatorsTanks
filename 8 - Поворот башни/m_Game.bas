Attribute VB_Name = "m_Game"
Public isStop As Boolean
Private tanks As Collection











Public Sub Game()

Dim tank As c_Tank
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
        For Each tank In tanks
            tank.Render
        Next tank

        DoEvents

        If isStop = True Then
            Exit Sub
        End If
        
    Loop
    


End Sub


Private Sub FindTanks()
Dim shp As Visio.Shape
Dim tank As c_Tank
    
    For Each shp In Application.ActivePage.Shapes
        If shp.CellExists("User.GameObject", 0) <> 0 Then
            If shp.Cells("User.GameObject").Result(visNumber) = 1 Then
                Set tank = New c_Tank
                Set tank.shp = shp
                tank.Activate
                tank.id = Str(tanks.Count)
                tanks.Add tank, tank.id
                Set tank.tanks = tanks
            End If
        End If
    Next shp
    
End Sub
