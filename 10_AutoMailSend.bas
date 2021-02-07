Attribute VB_Name = "AutoSendMail"
Sub Auto_Send_Mail()
       
    '�}�N�����׌v��
    Dim TIMEIN As Single
    Dim TIMEOUT As Single
    Dim TIMEDIFF As Single
    TIMEIN = Timer
    
    Dim inputPara(4) As String
    Dim outputPara() As String
    Dim mailBodyTitle As String
    Dim mailBodyFix As String
    
    Dim mailSendTo As String, mailSendCC As String, mailSubject As String, mailBody As String
    
    '��ʕ`��̒�~
    Call ScreenDrawStop(True)
    
    mailSubject = Sheets("���[�����e").Range("C2").Value
    mailBodyFix = Sheets("���[�����e").Range("C4").Value
    
    For Row = 4 To 203

'        inputPara(0) = Cells(Row, 3).Value
'        inputPara(1) = Cells(Row, 7).Value
'        inputPara(2) = Cells(Row, 8).Value
'        inputPara(3) = Cells(Row, 9).Value
'        inputPara(4) = Cells(Row, 13).Value
        
        '��Ж�
        inputPara(0) = Range("C" & Row).Value
        
        '������
        inputPara(1) = Range("G" & Row).Value
        
        '�S���Җ�
        inputPara(2) = Range("H" & Row).Value
        
        '���� �d�|���������A�h���X
        inputPara(3) = Range("I" & Row).Value
        
        '���񑗐M�v�ۃt���O
        inputPara(4) = Range("M" & Row).Value
        
        If inputPara(4) = "��" Then
        
            mailBodyTitle = inputPara(0) & " " & inputPara(1) & " " & inputPara(2) & "�l"
            mailBody = mailBodyTitle & vbCrLf & mailBodyFix
            
            mailSendTo = inputPara(3)
            mailSendCC = ""
            Call mail_send(mailSendTo, mailSendCC, mailSubject, mailBody)
            
        End If
        
        Debug.Print ("Row=" & Row)
    Next Row
    
    '��ʕ`��̍ĊJ
    Call ScreenDrawStop(False)
    
    TIMEOUT = Timer
    TIMEDIFF = TIMEOUT - TIMEIN
    
    MsgBox "�����l�ł���!" & vbCrLf & "�����ɂ����������Ԃ�" & Round(TIMEDIFF, 1) & "�b�ł��B"
    
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
            .To = mailSendTo                         '���[������
            .cc = mailSendCC                         '���[��CC
            .subject = mailSubject                   '���[������
            .body = mailBody                         '���[���{��
            .BodyFormat = olFormatPlain              '���[���̌`��
            .Display
            '.Send
        End With

End Function
