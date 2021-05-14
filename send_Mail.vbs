'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 4.0
'
' NAME: WeissAlex
'
' AUTHOR: SA , SB
' DATE  : 24.08.2007
'
' COMMENT: ������ ������������ ��� �������� ����� � ������������ ���� ���
'
'==========================================================================

Const cdoSendUsingPickup = 1
Const cdoSendUsingPort = 2 'Must use this to use Delivery Notification
Const cdoAnonymous = 0
Const cdoBasic = 1 ' clear text
Const cdoNTLM = 2 'NTLM
'Delivery Status Notifications
Const cdoDSNDefault = 0 'None
Const cdoDSNNever = 1 'None
Const cdoDSNFailure = 2 'Failure
Const cdoDSNSuccess = 4 'Success
Const cdoDSNDelay = 8 'Delay
Const cdoDSNSuccessFailOrDelay = 14 'Success, failure or delay

set objMsg = CreateObject("CDO.Message")
set objConf = CreateObject("CDO.Configuration")

Set objFlds = objConf.Fields
With objFlds

  .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.rambler.ru"
  .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
  .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "name"
  .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "login"
  .Update
End With

strBody = "� ���������" & VbCrLf
strBody = strBody & "��������� ������� ���������� " & VbCrLf
strBody = strBody & "��������� ������� �������������� ����������" & VbCrLf
strBody = strBody & "��� ��� ��� '�����-����'" & VbCrLf
strBody = strBody & "���. " & VbCrLf
strBody = strBody & "���. " & VbCrLf
strBody = strBody & "" & VbCrLf

With objMsg
  Set .Configuration = objConf
  .To = "test@mail.com"
  .From = "��� ���� " & """" & "����� ����" & """" 
  .Subject = "����� 652"
  '.TextBody = strBody
  .HTMLBody = strBody
  '.Addattachment "c:\temp\2FD.tmp"
  .Fields("urn:schemas:mailheader:disposition-notification-to") = "mail@mail.ru"
  .Fields("urn:schemas:mailheader:return-receipt-to") = "mail@mail.ru"
  .DSNOptions = cdoDSNSuccessFailOrDelay
  .Fields.update
  .Send
End With