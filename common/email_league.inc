<%
    '***********************************************************
	' email_league()
	' - Set the email_message variable before calling this include.
	'
	'   <Called By>
	'   * tradeoffers.asp
	'   * transelect.asp
	'   * xxxxx.asp
	'   * xxxxxx.asp
	'***********************************************************
	
	'Response.Write "TOP of email_league.inc <br>"
	if ISNULL(wEmailOwnerID) then
	   objEmail.Open  "SELECT * from tblOwners " &_
		  		       "WHERE ownerid <> 99 and "&wAlert&" = TRUE and (receiveTexts = TRUE or receiveEmails = TRUE) ORDER by OwnerID ", objConn,3,3,1
	else
	   objEmail.Open  "SELECT * from tblOwners " &_
		  	          "WHERE ownerid = "&wEmailOwnerID&" and "&wAlert&" = TRUE and (receiveTexts = TRUE or receiveEmails = TRUE) ", objConn,3,3,1					   
	end if
	
    if objEmail.Recordcount > 0 then	
		wAppend = 0       
		if objEmail.Fields("receiveEmails").Value = TRUE then
		   wAppend = 1
		   email_to = objEmail.Fields("HomeEmail").Value  
		end if
		
		 if objEmail.Fields("receiveTexts").Value = TRUE then		   
		   if wAppend = 1 then
		      email_to = email_to & "," & objEmail.Fields("textMessages").Value   
		   else		      
			  email_to = objEmail.Fields("textMessages").Value
		   end if
		end if
						
		objEmail.MoveNext
				
		While Not objEmail.EOF	
		   if objEmail.Fields("receiveEmails").Value = TRUE then		   
			  email_to = email_to & "," & objEmail.Fields("HomeEmail").Value
		   end if
			
			 if objEmail.Fields("receiveTexts").Value = TRUE then		   
			  email_to = email_to & "," & objEmail.Fields("textMessages").Value
		   end if

 		   objEmail.MoveNext
		Wend
		objEmail.Close
				
		host         = "win-mail06.hostmanagement.net"   'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite)
		username     = "igbl_commish@igbl.org" 'A valid email address you have setup
		from_address = "igbl_commish@igbl.org" 'if your mail is hosted with HostMySite this has to match the email address above
		password     = "SpursRtheChamps2014" 'Password for the above email address
		reply_to     = "noReply@igbl.org"  'Enter the email you want customers to reply to
		port         = "25" 'This is the default port. Try port 25, 50, 587 if this port gives you issues and your mail is hosted with HostMySite

		Set ObjSendMail = CreateObject("CDO.Message")
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = host
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
		ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
		ObjSendMail.Configuration.Fields.Update
		
		'Response.Write "email_to = " & email_to  & "<br>"
		
		ObjSendMail.To       = email_to
		ObjSendMail.Subject  = email_subject
		ObjSendMail.From     = from_address
		ObjSendMail.HTMLBody = email_message
	

		ObjSendMail.Send
		set ObjSendMail = Nothing
    else
	    objEmail.Close
	end if
%>