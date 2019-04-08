Attribute VB_Name = "tools"
Public Const rad360 = 6.2831853071796
Public Const rad180 = 3.1415926535898
Public Const rad90 = 1.5707963267949


Public Function GetCurrentPosition(ByRef shp As Visio.Shape) As c_Point
Dim pnt As c_Point
Set pnt = New c_Point
    
    pnt.x = shp.Cells("PinX").Result(visInches)
    pnt.y = shp.Cells("PinY").Result(visInches)
    
    Set GetCurrentPosition = pnt
End Function
