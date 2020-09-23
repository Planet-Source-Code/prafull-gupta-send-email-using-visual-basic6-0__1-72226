<div align="center">

## Send email using Visual Basic6\.0


</div>

### Description

This article describes how to pro grammatically send an e-mail message from Visual Basic by using the Collaboration Data Objects (CDO 1.x) Libraries. This example need the CDO 1.x was correctly installed on your computer.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Prafull Gupta](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/prafull-gupta.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/prafull-gupta-send-email-using-visual-basic6-0__1-72226/archive/master.zip)





### Source Code

Send Email using Visual Basic6.0<br>
<br><p>This article describes how to pro grammatically send an e-mail message from Visual Basic by using the Collaboration Data Objects (CDO 1.x) Libraries. This example need the CDO 1.x was correctly installed on your computer. You can create programmable messaging objects, and then use their properties and methods to meet the needs of your application.
The sample code can be used from Visual Basic for Applications (VBA), Project to send e-mail messages through the CDO 1.x Library.<p>
<br>
Dim file As String<br>
‘ App.path takes the application path and after that user have to provide the file name like as in given code.
<br>
file = App.Path & “filename.ext”
<br>
‘For creating new Message object
<br>
Dim iMsg As New CDO.Message <br>
Dim iDsrc As CDO.IDataSource <br>
Set iDsrc = iMsg ' (QueryInterface)<br>
Dim iConf As New CDO.Configuration <br>
Dim Flds As Variant<br>
Set Flds = iConf.Fields<br>
With Flds<br>
.Item(cdoSendUsingMethod) = cdoSendUsingPort<br>
.Item(cdoSMTPServer) = "smtp.gmail.com"
'"smtp.myServer.com"
<br>
.Item(cdoSMTPServerPort) = "25"  ‘port no <br>
.Item(cdoSMTPConnectionTimeout) = 1000 ' quick timeout
<br>
.Item(cdoSMTPAuthenticate) = cdoBasic <br>
.Item(cdoSMTPUseSSL) = True <br>
.Item(cdoSendUserName) = "abc" ‘”username” <br>
.Item(cdoSendPassword) = "***" ‘”password” <br>
.Update <br>
End With <br>
With iMsg <br>
Set .Configuration = iConf <br>
.To = "abc@gmail.com" <br>
.From = "test@gmail.com" <br>
.Subject = "Test” ‘write the subject line here for your mail
<br>
.TextBody = "This is the test body” ‘write the body part here
<br>
.AddAttachment file ‘used for attachement.You can attach as many files
<br>
.send
<br>
End With
<br>

