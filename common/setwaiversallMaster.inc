<%

	Set SwaWaivers     = Server.CreateObject("ADODB.RecordSet")
	Set SwaLineups 	   = Server.CreateObject("ADODB.RecordSet")
	Set SwaNextRun	   = Server.CreateObject("ADODB.RecordSet")
	Set SwaStagger	   = Server.CreateObject("ADODB.RecordSet")

	'*****************************************************
	'Run setwaiversall event if it hasn't been run today.
	'*****************************************************
	'Response.Write "Top of setwaiversall.inc <br>"
	SwaWaivers.Open "SELECT * FROM tbltimedEvents where event = 'setwaiversall' and nextrun_EST < now() ", objConn,3,3,1

	if SwaWaivers.Recordcount > 0 then
	
		'********************************************************************
		'Check the time that this action is being triggered.  Either before or after
		'midnight.  If it is before midnight then next actions will be based on tomorrow.   
		'If after midnight then next actions will be based on today's date.		
		'If tomorrow is a game, then set the pendingwaivers date = to the cutofftimestagger.
		'Note that that code adds 1 hour to the time because the times
		'in the database are CST but the server is hosted on EST.  If tomorrow is
		'not a game day then set the nextrun date to be 7:00 PM EST.
		'*********************************************************************		
		
		wTime= time()
		wpos = Instr(wTime, "PM")
	  
	    if wpos > 0 then   'Before Midnight
		   dlineups = date()
           SwaNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = (date() + 1) ", objConn,3,3,1
		   strSQL = "update tblPlayers " & _
		            "SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() + 1 " & _
		            "WHERE playerStatus = 'F' or playerStatus = 'S' or RentalPlayer = 1 "
		else
		   dlineups = date() - 1
		   SwaNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = date() ", objConn,3,3,1		
		   strSQL = "update tblPlayers " & _
		             "SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() " & _
		             "WHERE playerStatus = 'F' or playerStatus = 'S' or RentalPlayer = 1 "
	    end if
				
		objConn.Execute strSQL

		
		if SwaNextRun.Recordcount > 0 then
			dnextrun = SwaNextRun.Fields("cutofftimestagger").Value + 1/24
		else
			if wpos > 0 then
			 dnextrun = date() + 1 + (19/24)
			else
			 dnextrun = date() + (19/24)			
			end if
		end if

		strSQL = "update tbltimedEvents SET lastrun_EST = now(), nextrun_EST = '"&dnextrun&"' WHERE event = 'setwaiversall' "
		objConn.Execute strSQL
		SwaNextRun.Close
		
		SwaStagger.Open "SELECT * From tblparameterCtl where param_indicator = TRUE and param_name = 'STAGGER_WINDOW' ", objConn,3,3,1
		
		if SwaStagger.Recordcount > 0 then		
		   strSQL = "update tblParameterCtl set param_indicator = 0 where param_name = 'STAGGER_WINDOW' "
		   objConn.Execute strSQL				   
		
		   'Response.Write "Stagger Window Close Email<br>"
		   '**************************************************************************
		   'Send Email Notification to the League that we are in the Stagger Period!
		   '**************************************************************************

		   email_to     = "keith_dlott@igbl.org, dennis_myers@igbl.org, chris_Jones@igbl.org, gary_rothballer@igbl.org, antone_adams@igbl.org, jack_white@igbl.org, jeff_peskin@igbl.org, fred_curry@igbl.org, anthony_white@igbl.org, cliff_fox@igbl.org, lamont_davis@igbl.org"  'Enter the email you want to send the form to
		   email_subject= "Stagger Window Closed!"  'You can put whatever subject here
		   host         = "mail.igbl.org"   'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite) 
		   username     = "igbl_commish@igbl.org"  'A valid email address you have setup 
		   from_address ="igbl_commish@igbl.org" 'If your mail is hosted with HostMySite this has to match the email address above 
		   password     = "SpursRtheChamps2014" 'Password for the above email address

		   reply_to     = "noReply@igbl.org"  'Enter the email you want customers to reply to
		   port = "25" 'This is the default port. Try port 50 if this port gives you issues and your mail is hosted with HostMySite
		
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
		
		   email_message       = "The Stagger Window Period has Ended.<br /><br /> All ***ACTIVE PLAYERS IN YOUR LINEUP*** are now available to be traded or waived!<br /><br />"
		
		   ObjSendMail.To      = email_to
		   ObjSendMail.Subject = email_subject
		   ObjSendMail.From    = from_address    
		   ObjSendMail.HTMLBody= email_message   
		   ObjSendMail.Send
		   set ObjSendMail     = Nothing
		
		   '********************************************************************************
		   'End of Send Email Notification to the League that we are in the Stagger Period!
		   '********************************************************************************		
		end if
		
		SwaStagger.Close
		'********************************************************************************************
		'If today was a game day then set final_lineup flag and move lineups to tbl_lineups_staggered
		'********************************************************************************************
		SwaLineups.Open	"SELECT * FROM tbl_lineups where gameday = #"&dlineups&"# ", objConn,3,3,1
		if SwaLineups.Recordcount > 0 then
			strSQL = "update tbl_lineups set final_lineup = 1 where gameday = #"&dlineups&"# "
			objConn.Execute strSQL
			
			strSQL = "delete from tbl_lineups_staggered"
			objConn.Execute strSQL
			
			strSQL = "insert into tbl_lineups_staggered select * from tbl_lineups where gameday = #"&dlineups&"# "
			objConn.Execute strSQL						
		end if
		
		SwaLineups.Close
	end if
	
	SwaWaivers.Close


%>