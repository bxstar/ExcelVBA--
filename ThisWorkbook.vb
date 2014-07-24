Private Sub Workbook_Open()
'定义工作簿是否被修改
    
    'Daily的源数据表是否被修改
    Dim isChangeDaily As Boolean
    'Campaign的源数据表是否被修改
    Dim isChangeCampaign As Boolean
    'Channel对应表是否被修改
    Dim isChangeChannelRelation As Boolean
    '配置表是否被修改
    Dim isChangeSetUp As Boolean
    
    isChangeDaily = FunIsChangeDaily()
    isChangeCampaign = FunIsChangeCampaign()
    isChangeChannelRelation = FunIsChangeChannelRelation()
    isChangeSetUp = FunIsChangeSetUp()
    
    If isChangeDaily Then
        
'        MsgBox "Daily被修改"
        '生成Daily报表
        Call WriteDailyReport
        Call WriteChannelReport
        Sheets("tmp").Range("I2") = 0
        Sheets("tmp").Range("I4") = 0
    Else
        If isChangeChannelRelation Then
        'Channel对应表被修改
            Call WriteChannelReport
            Sheets("tmp").Range("I4") = 0
        End If
    End If
    
    If isChangeCampaign Then
        
'        MsgBox "Campaign被修改"
        '生成Campaign报表
        Call WriteCampaignReport
        Sheets("tmp").Range("I3") = 0
    End If
    
    If isChangeSetUp Then
        '设置数据指标的显示
        Call SetMetricDataDisplay("Daily")
        Call SetMetricDataDisplay("Campaign")
        Sheets("tmp").Range("I5") = 0
    End If
    
    '如果有改变则保存更改
    If isChangeCampaign Or isChangeDaily Or isChangeSetUp Then
        ThisWorkbook.Save
    Else
'        MsgBox "没有内容被修改"
    End If


End Sub


Function FunIsChangeDaily() As Boolean
'判断Daily元数据表是否被修改
    If Sheets("tmp").Range("I2") = 1 Then
        FunIsChangeDaily = True
    Else
        FunIsChangeDaily = False
    End If
    

End Function


Function FunIsChangeCampaign() As Boolean
'判断Campaign元数据表是否被修改
    If Sheets("tmp").Range("I3") = 1 Then
        FunIsChangeCampaign = True
    Else
        FunIsChangeCampaign = False
    End If

End Function

Function FunIsChangeChannelRelation() As Boolean
'Channel对应表是否被修改
    If Sheets("tmp").Range("I4") = 1 Then
        FunIsChangeChannelRelation = True
    Else
        FunIsChangeChannelRelation = False
    End If

End Function

Function FunIsChangeSetUp() As Boolean
'配置表是否被修改
    If Sheets("tmp").Range("I5") = 1 Then
        FunIsChangeSetUp = True
    Else
        FunIsChangeSetUp = False
    End If

End Function

