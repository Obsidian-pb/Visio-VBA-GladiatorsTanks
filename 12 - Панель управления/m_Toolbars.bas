Attribute VB_Name = "m_toolbars"
Sub AddToolBar()

    Dim Bar As CommandBar

    Set Bar = Application.CommandBars.Add(Position:=msoBarTop, Temporary:=True)
    With Bar
        .Name = "GladiatorsTanks"
        .Visible = True
    End With
    
    AddButtons

End Sub

Private Sub AddButtons()

    Dim Bar As CommandBar
    Dim Button As CommandBarButton

    Set Bar = Application.CommandBars("GladiatorsTanks")
    
    '---Кнопка Старт
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Старт"
        .Tag = "Start"
        .OnAction = "Game"
        .TooltipText = "Запустить игру"
        .FaceID = 186
    End With
    
    '---Кнопка Стоп
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Стоп"
        .Tag = "Stop"
        .OnAction = "StopGame"
        .TooltipText = "Остановить игру"
        .FaceID = 228
    End With
    
    '---Кнопка Очистить
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "Очистить"
        .Tag = "Clear"
        .OnAction = "ClearShells"
        .TooltipText = "Очистить все снаряды"
        .FaceID = 1564
    End With
    
    Set Button = Nothing
    
    
    
End Sub
