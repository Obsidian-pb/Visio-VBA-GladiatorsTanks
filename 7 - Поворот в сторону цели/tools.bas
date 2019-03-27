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

Public Function GetAngleBetweenPoints(ByRef pnt1 As c_Point, ByRef pnt2 As c_Point) As Double
Dim dx As Double
Dim dy As Double
    
    dx = pnt2.x - pnt1.x
    dy = pnt2.y - pnt1.y
    GetAngleBetweenPoints = GetATAN(dx, dy)
    
End Function

Public Function GetATAN(ByVal x As Double, ByVal y As Double) As Double
Dim val As Double
    
    If x = 0 Then
        If y > 0 Then
            GetATAN = rad90
        Else
            GetATAN = rad180 + rad90
        End If
        Exit Function
    End If
    
    val = Atn(y / x)  '-y/x = y/-x,
    
    If x > 0 And y > 0 Then
        GetATAN = val
    ElseIf x < 0 And y > 0 Then
        GetATAN = val + rad180
    ElseIf x < 0 And y < 0 Then
        GetATAN = val + rad180
    ElseIf x > 0 And y < 0 Then
        GetATAN = val + rad360
    End If
End Function

Public Function IsRight(ByVal direction As Double, ByVal targetDirection As Double) As Boolean
    
    If GetAngleBetweenAngles(direction, targetDirection) < 0 Then
        IsRight = False
    Else
        IsRight = True
    End If
    
End Function

Public Function GetAngleBetweenAngles(ByVal direction As Double, ByVal targetDirection As Double) As Double
Dim diff As Double
    diff = direction - targetDirection
    GetAngleBetweenAngles = diff
    If diff < -rad180 Then
        GetAngleBetweenAngles = diff + rad360
    ElseIf diff > rad180 Then
        GetAngleBetweenAngles = diff - rad360
    End If
End Function
