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
    
    '---������ �����
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "�����"
        .Tag = "Start"
        .OnAction = "Game"
        .TooltipText = "��������� ����"
        .FaceID = 186
    End With
    
    '---������ ����
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "����"
        .Tag = "Stop"
        .OnAction = "StopGame"
        .TooltipText = "���������� ����"
        .FaceID = 228
    End With
    
    '---������ ��������
    Set Button = Bar.Controls.Add(Type:=msoControlButton)
    With Button
        .Caption = "��������"
        .Tag = "Clear"
        .OnAction = "ClearShells"
        .TooltipText = "�������� ��� �������"
        .FaceID = 1564
    End With
    
    Set Button = Nothing
    
    
    
End Sub
