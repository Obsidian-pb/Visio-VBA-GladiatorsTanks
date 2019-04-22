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
    bodyRotationSpeed = shp.Cells("Prop.RotationSpeed").Result(visNumber)

    turretDirection = direction + shp.Shapes(2).Cells("Angle").Result(visRadians)
    turretRotationSpeed = shp.Cells("Prop.TurretRotationSpeed").Result(visNumber)
    shotSpeed = shp.Cells("Prop.ShotSpeed").Result(visNumber)
End Sub

Public Sub Render()
    
Dim shell As c_Shell
    
    
    If tanks.Count = 1 Then Exit Sub
    
    'Наводимся на цель
    If target Is Nothing Then
        'Наводимся на цель
        Do
            Set target = tanks(Int(Rnd() * tanks.Count) + 1)
            If target.id <> id Then Exit Do
        Loop
    End If
    
    'Движение
    nextPoisition.x = curPoisition.x + bodySpeed * Cos(direction)
    nextPoisition.y = curPoisition.y + bodySpeed * Sin(direction)
    
    shp.Cells("PinX").Formula = nextPoisition.x
    shp.Cells("PinY").Formula = nextPoisition.y
    Set curPoisition = nextPoisition
    
    'Поворачиваем корпус
    targetDirection = GetAngleBetweenPoints(nextPoisition, target.curPoisition)
    If IsRight(direction, targetDirection) Then
        direction = direction - bodyRotationSpeed
    Else
        direction = direction + bodyRotationSpeed
    End If
    shp.Cells("Angle").Formula = direction - rad90
    
    'Поворачиваем башню в сторону цели
    If IsRight(turretDirection, targetDirection) Then
        turretDirection = turretDirection - turretRotationSpeed
    Else
        turretDirection = turretDirection + turretRotationSpeed
    End If
    If turretDirection > rad360 Then turretDirection = turretDirection - rad360
    If turretDirection < 0 Then turretDirection = turretDirection + rad360
    shp.Shapes(2).Cells("Angle").Formula = turretDirection - direction
    
    
    
    'Стряляем
    If i > shotSpeed And IsSameAngle(turretDirection, targetDirection, 2) Then
        Set shellStart = GetShellStartPosition(shp.Shapes(2))
        Set shell = New c_Shell
        Set shell.shp = Application.ActiveWindow.Page.DrawOval(shellStart.x - shellRadiuss, shellStart.y - shellRadiuss, _
                                                shellStart.x + shellRadiuss, shellStart.y + shellRadiuss)
        Set shell.currentPosition = shellStart
        shell.direction = targetDirection + (Rnd() * 0.2 - 0.1)
        shell.id = id & "_" & shotNumber
        Set shell.firedTank = Me
        shell.SetLayer "Shells"
        Set shell.tanks = tanks
        shells.Add shell, shell.id

        shotNumber = shotNumber + 1
        i = 0
    End If
    
    
    i = i + 1
End Sub


Public Sub RizeScore()
On Error Resume Next
    scores = scores + 1
    shp.Cells("Prop.DestroyedTanks").Formula = scores
End Sub

