VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True






Public Sub Game()

Dim shp As Visio.Shape
Dim speed As Double
Dim x As Double
Dim y As Double
Dim endTime As Date



    endTime = DateAdd("s", 3, Now())
    
    Set shp = Application.ActivePage.Shapes(1)
    
    x = shp.Cells("PinX").Result(visInches)
    y = shp.Cells("PinY").Result(visInches)
    speed = shp.Cells("Prop.Speed").Result(visNumber)
    
    
    Do While endTime > Now()
        x = x + speed
        y = y + speed
        shp.Cells("PinX").Formula = x
        shp.Cells("PinY").Formula = y





        DoEvents
        
        

        
    Loop
    
    
    
    
    
    







End Sub


