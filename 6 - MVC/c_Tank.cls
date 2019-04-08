VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Tank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public shp As Visio.Shape
Public id As String

Public direction As Double
Public curPoisition As c_Point
Public nextPoisition As c_Point
Public bodySpeed As Double
Public bodyRotationSpeed As Double
Public scores As Integer

Public tanks As Collection
Public target As c_Tank
Public targetDirection As Double

Public turretRotationSpeed As Double
Public turretDirection As Double
Public shotSpeed As Integer
Public shells As Collection
Public shellStart As c_Point


Private i As Integer
Private shotNumber As Long


Public Sub Activate()
    Set nextPoisition = New c_Point
    Set curPoisition = GetCurrentPosition(shp)
    direction = shp.Cells("Angle").Result(visRadians) + rad90
    bodySpeed = shp.Cells("Prop.Speed").Result(visNumber)
    
    
End Sub

Public Sub Render()
    
    'Движение
    nextPoisition.x = curPoisition.x + bodySpeed * Cos(direction)
    nextPoisition.y = curPoisition.y + bodySpeed * Sin(direction)
    
    shp.Cells("PinX").Formula = nextPoisition.x
    shp.Cells("PinY").Formula = nextPoisition.y
    Set curPoisition = nextPoisition
    
    i = i + 1
End Sub
