Attribute VB_Name = "m_Game"
Public isStop As Boolean
Private tanks As Collection
Private shells As Collection










Public Sub Game()

Dim tank As c_Tank
Dim shell As c_Shell
Dim speed As Double
Dim x As Double
Dim y As Double
Dim direction As Double
Dim endTime As Date

    isStop = False
    endTime = DateAdd("s", 60, Now())
    
    Set shells = New Collection
    Set tanks = New Collection
    FindTanks
    

    Do While endTime > Now()
        If tanks.Count > 0 Then
            For Each tank In tanks
                tank.Render
            Next tank
        End If
        
        If shells.Count > 0 Then
            For Each shell In shells
                If shell.Render Then
                    shells.Remove shell.id
                    shell.shp.Delete
                    Set shell = Nothing
                End If
            Next shell
        End If
        
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
                Set tank.shells = shells
                tank.Activate
                tank.id = Str(tanks.Count)
                tanks.Add tank, tank.id
                Set tank.tanks = tanks
            End If
        End If
    Next shp
    
End Sub

Public Sub ClearShells()
Dim vsoSelection As Visio.Selection
    Set vsoSelection = Application.ActiveWindow.Page.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Shells")
    vsoSelection.Delete
End Sub
