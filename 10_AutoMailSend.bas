Attribute VB_Name = "AutoSendMail"
Sub Auto_Send_Mail()
       
    'マクロ負荷計測
    Dim TIMEIN As Single
    Dim TIMEOUT As Single
    Dim TIMEDIFF As Single
    TIMEIN = Timer
    
    Dim inputPara(4) As String
    Dim outputPara() As String
    Dim mailBodyTitle As String
    Dim mailBodyFix As String
    
    Dim mailSendTo As String, mailSendCC As String, mailSubject As String, mailBody As String
    
    '画面描画の停止
    Call ScreenDrawStop(True)
    
    mailSubject = Sheets("メール内容").Range("C2").Value
    mailBodyFix = Sheets("メール内容").Range("C4").Value
    
    For Row = 4 To 203

'        inputPara(0) = Cells(Row, 3).Value
'        inputPara(1) = Cells(Row, 7).Value
'        inputPara(2) = Cells(Row, 8).Value
'        inputPara(3) = Cells(Row, 9).Value
'        inputPara(4) = Cells(Row, 13).Value
        
        '会社名
        inputPara(0) = Range("C" & Row).Value
        
        '部署名
        inputPara(1) = Range("G" & Row).Value
        
        '担当者名
        inputPara(2) = Range("H" & Row).Value
        
        '宛先 Ｅ－ｍａｉｌアドレス
        inputPara(3) = Range("I" & Row).Value
        
        '今回送信要否フラグ
        inputPara(4) = Range("M" & Row).Value
        
        If inputPara(4) = "○" Then
        
            mailBodyTitle = inputPara(0) & " " & inputPara(1) & " " & inputPara(2) & "様"
            mailBody = mailBodyTitle & vbCrLf & mailBodyFix
            
            mailSendTo = inputPara(3)
            mailSendCC = ""
            Call mail_send(mailSendTo, mailSendCC, mailSubject, mailBody)
            
        End If
        
        Debug.Print ("Row=" & Row)
    Next Row
    
    '画面描画の再開
    Call ScreenDrawStop(False)
    
    TIMEOUT = Timer
    TIMEDIFF = TIMEOUT - TIMEIN
    
    MsgBox "お疲れ様でした!" & vbCrLf & "処理にかかった時間は" & Round(TIMEDIFF, 1) & "秒です。"
    
End Sub

Sub ScreenDrawStop(ByVal Flag As Boolean)

    With Application
        .EnableEvents = Not Flag
        .ScreenUpdating = Not Flag
        .DisplayAlerts = Not Flag
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)
    End With
    
End Sub

Function mail_send(mailSendTo As String, mailSendCC As String, mailSubject As String, mailBody As String)

        Dim objOutlook As Outlook.Application
        Set objOutlook = New Outlook.Application
        
        Dim objMailItem As Outlook.MailItem
        Set objMailItem = objOutlook.CreateItem(olMailItem)
        
        With objMailItem
            .To = mailSendTo                         'メール宛先
            .cc = mailSendCC                         'メールCC
            .subject = mailSubject                   'メール件名
            .body = mailBody                         'メール本文
            .BodyFormat = olFormatPlain              'メールの形式
            .Display
            '.Send
        End With

End Function
