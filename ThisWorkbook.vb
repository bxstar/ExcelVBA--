Private Sub Workbook_Open()
'定义工作簿是否被修改
    Dim isChangeDaily As Boolean, isChangeCampaign As Boolean
    
    isChangeDaily = FunIsChangeDaily()
    isChangeCampaign = FunIsChangeCampaign()
    
    If isChangeDaily Then
        
'        MsgBox "Daily被修改"
        '生成Daily报表
        Call WriteDailyReport
        Call WriteChannelReport
        Sheets("tmp").Range("I2") = 0
    End If
    
    If isChangeCampaign Then
        
'        MsgBox "Campaign被修改"
        '生成Campaign报表
        Call WriteCampaignReport
        Sheets("tmp").Range("I3") = 0
    End If
    
    '如果有改变则保存更改
    If isChangeCampaign Or isChangeDaily Then
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
