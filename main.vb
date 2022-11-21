Sub Create_Bar()
    Dim cBar As CommandBar, ctlBtt As CommandBarButton
    Delete_Bar
    Set cBar = Application.CommandBars.Add("Kieu Manh_Bar")
    With cBar
        .Position = msoBarTop
        .Visible = True
        Set ctlBtt = .Controls.Add(msoControlButton)
        Set ctlBtt1 = .Controls.Add(msoControlButton)
        Set ctlBtt2 = .Controls.Add(msoControlButton)
        
    End With
    With ctlBtt
        .Caption = "Delete nude"
        .Style = msoButtonIconAndCaption
        .OnAction = "MainDeleteNull"
        .FaceId = 107
    End With
    With ctlBtt1
        .Caption = "MainAll"
        .Style = msoButtonIconAndCaption
        .OnAction = "MainAll"
        .FaceId = 108
    End With
    With ctlBtt2
        .Caption = "TierAll"
        .Style = msoButtonIconAndCaption
        .OnAction = "TierAll"
        .FaceId = 109
    End With
End Sub
