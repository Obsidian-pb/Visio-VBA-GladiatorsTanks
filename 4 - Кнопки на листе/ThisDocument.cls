VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private isStop As Boolean



Private Sub CommandButton1_Click()
    Game
End Sub


Private Sub CommandButton2_Click()
    isStop = True
End Sub





Public Sub Game()

Dim shp As Visio.Shape
Dim speed As Double
Dim x As Double
Dim y As Double
Dim direction As Double
Dim endTime As Date

    isStop = False

    endTime = DateAdd("s", 60, Now())
    
    Set shp = Application.ActivePage.Shapes(1)
    
    x = shp.Cells("PinX").Result(visInches)
    y = shp.Cells("PinY").Result(visInches)
    speed = shp.Cells("Prop.Speed").Result(visNumber)
    direction = shp.Cells("Angle").Result(visRadians) + rad90
    
    
    
    Do While endTime > Now()
        x = x + speed * Cos(direction)
        y = y + speed * Sin(direction)
        shp.Cells("PinX").Formula = x
        shp.Cells("PinY").Formula = y





        DoEvents
        
        
        If isStop = True Then
            Exit Sub
        End If
        
    Loop
    
    
    
    
    
    







End Sub


