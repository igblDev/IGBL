<%	
	Set SsaStagger         = Server.CreateObject("ADODB.RecordSet")
	Set SsaNextRun	       = Server.CreateObject("ADODB.RecordSet")
	Set SsaEmail	       = Server.CreateObject("ADODB.RecordSet")
		
    '*******************************************************
	'Run setStaggeredAll event if it hasn't been run today.
	'*******************************************************
	'Response.Write "Top of setStageredAll.inc <br>"
	SsaStagger.Open "SELECT * FROM tbltimedEvents where event = 'setStaggeredAll' and nextrun_EST < now() ", objConn,3,3,1
					 
	if  SsaStagger.Recordcount > 0 then	

		strSQL = "update tblPlayers SET playerStatus = 'S' WHERE playerStatus = 'F' "
		objConn.Execute strSQL
	
		'***********************************************************************************
		'Check the time that this action is being triggered.  Either before or after
		'midnight.  If it is before midnight then next actions will be based on tomorrow.   
		'If after midnight then next actions will be based on today's date.
		'Set the time for the next setStaggeredAll run.  The next run will 
		'be the cutoff time of thee next game day. 
		'Note that that code adds 1 hour to the time because the times
		'in the database are CST but the server is hosted on EST.   
		'***********************************************************************************	
		wTime = time()
	    wpos = Instr(wTime, "PM")
	  
	    if wpos > 0 then   'Before Midnight
	       SsaNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday >= (date() + 1) ", objConn,3,3,1
  	    else
	       SsaNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday >= date() ", objConn,3,3,1
	    end if
		
		'***************************************************
		' If no records exist then the season is over.
		'***************************************************
		if SsaNextRun.Recordcount > 0 then
			dnextrun = SsaNextRun.Fields("cutofftime").Value + 1/24
		else
			dnextrun = now() + 365
		end if
		
		strSQL = "update tbltimedEvents SET lastrun_EST = now(), nextrun_EST = '"&dnextrun&"' WHERE event = 'setStaggeredAll' "
		objConn.Execute strSQL
		
		strSQL = "update tblParameterCtl set param_indicator = 1 where param_name = 'STAGGER_WINDOW' "
		objConn.Execute strSQL
		
		SsaNextRun.Close
		
		'**************************************************************************
		'Send Email Notification to the League that we are in the Stagger Period!
		'**************************************************************************		
		wEmailOwnerID       = null
		wAlert              = "receiveStaggerAlerts"
		email_subject       = "Stagger Window Open!" 
		email_message       = "*** Active Players in Your Lineup are Locked from Activity! ***<br>"
%>
		<!--#include virtual="Common/email_league.inc"-->
				   
<%
	end if  'End if SsaStagger.Recordcount > 0
	
	SsaStagger.Close
	

	
%>