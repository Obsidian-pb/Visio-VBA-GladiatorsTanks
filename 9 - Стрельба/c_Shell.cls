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
'Private i As Long



Public Function Render() As Boolean
    currentPosition.x = currentPosition.x + speed * Cos(direction)
    currentPosition.y = currentPosition.y + speed * Sin(direction)
    
    shp.Cells("PinX").Formula = currentPosition.x
    shp.Cells("PinY").Formula = currentPosition.y
        
    
End Function

