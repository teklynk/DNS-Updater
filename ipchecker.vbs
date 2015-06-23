IPCheckerURL = "http://icanhazip.com"
dyndnsURL = "https://dynamic.zoneedit.com/auth/dynamic.html?host=mydomainame1.com,mydomainame2.com&dnsto="
dydnsUername = "xxxx"
dydnsPassword = "xxxx"
PathToIPFile = "C:\scripts\myipaddress.txt"

'default values
myIPAddress = "0.0.0.0"
serverName = "WEB SERVER"
IPAlertStatus = "IP Address UNKNOWN"
alertStatus = "Not able to get IP"
serverDown=true
newIP=0

on error resume next 'skip errors

set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.Open "GET", IPCheckerURL, False
		xmlhttp.Send
		status = xmlhttp.status
		myIPAddress = xmlhttp.responseText

		if err.number <> 0 or status <> 200 then
			alertStatus = serverName&" - "&IPCheckerURL&" - Internet is down!? "
			myIPAddress
			serverDown
			ResetModem
			'ResetRouter
		else
			alertStatus = serverName&" - "&IPCheckerURL&" - Was reachable! "
			myIPAddress = myIPAddress
			serverDown=false
		end if
		
		ReadIPFile
		
		if newIP <> 0 then
			xmlhttp.Open "GET", dyndnsURL&myIPAddress, False, ""&dydnsUername&"", ""&dydnsPassword&""
			xmlhttp.Send
			status = xmlhttp.status
			DNSstatus = xmlhttp.responseText
		else
			DNSStatus = "No DNS Update Required "
		end if
		
		SendMail 'send an email

set xmlhttp = nothing

Sub ReadIPFile
	Set fs = CreateObject("Scripting.FileSystemObject")
	set t=fs.OpenTextFile(PathToIPFile,1,false)
		IPAddressOnFile=t.ReadAll
		t.close

		if StrComp(IPAddressOnFile,myIPAddress)<>0 then
			newIP = -1
			WriteToIPFile
			IPAlertStatus = "Your IP Address has changed! "
		else
			newIP = 0
			IPAlertStatus = "IP Address is OK "
		end if

	Set fs = Nothing
	Set t = Nothing
End Sub

Sub WriteToIPFile
    Set fs = CreateObject("Scripting.FileSystemObject")
	'open for writing
    Set f = fs.OpenTextFile(PathToIPFile,2,true)
		f.Write myIPAddress
		f.Close
	Set fs = Nothing
	Set f = Nothing
End Sub

Sub ResetModem
	xmlhttp.Open "GET", "http://192.168.100.1/reset.htm?reset_modem=Restart+Cable+Modem", False
	xmlhttp.Send
End Sub

'This does not work
Sub ResetRouter
	xmlhttp.Open "GET", "http://192.168.0.1/Management.asp?Reboot", False, "xxxx", "xxxx"
	xmlhttp.Send
End Sub

Sub SendMail
	EmailFrom = "noreply@gmail.com"
	EmailFromName = serverName&" SERVER ALERT"	
	EmailTo = "youremail@gmail.com"
	SMTPServer = "smtp.gmail.com"
	SMTPLogon = "youremail@gmail.com"
	SMTPPassword = "xxxx"
	SMTPSSL = True
	SMTPPort = 465
	cdoSendUsingPickup = 1    'Send message using local SMTP service pickup directory.
	cdoSendUsingPort = 2  'Send the message using SMTP over TCP/IP networking.
	cdoAnonymous = 0  ' No authentication
	cdoBasic = 1  ' BASIC clear text authentication
	cdoNTLM = 2   ' NTLM, Microsoft proprietary authentication

	EmailSubject = serverName&" SERVER ALERT"
	EmailBody = alertStatus&" - "&IPalertStatus&" - "&myIPAddress&" - "&DNSstatus&" - Date/Time: "&now()

	Set objMessage = CreateObject("CDO.Message")
		objMessage.Subject = EmailSubject
		objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
		objMessage.To = EmailTo
		objMessage.TextBody = EmailBody 
	
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		objMessage.Configuration.Fields.Update
		objMessage.Send
	Set objMessage = nothing
End Sub
