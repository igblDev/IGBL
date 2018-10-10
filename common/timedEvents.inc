<%
 	Dim objRSwaivers,strSQL,iPlayerClaimed,objRSTxns,objRejectWaivers,iPlayerWaived,iOwner,w_action
	
	Set objRSwaivers       = Server.CreateObject("ADODB.RecordSet")
	Set objStagger         = Server.CreateObject("ADODB.RecordSet")
	Set objRSTxns 	       = Server.CreateObject("ADODB.RecordSet")
	Set objRejectWaivers   = Server.CreateObject("ADODB.RecordSet")
	Set objNextRun	       = Server.CreateObject("ADODB.RecordSet")
	
	'********************************************************
	'Run pendingwaiversall event if it has not been run today.
	'********************************************************
	'Response.Write "Top of timedEvents.inc<br>"
	objRSwaivers.Open "SELECT * FROM tbltimedEvents " & _
	                  "where event = 'pendingwaiversall' and nextrun < now() ", cn,3,3,1

	if  objRSwaivers.Recordcount > 0 then
			objRSTxns.Open		"SELECT * FROM qryUpdatewaiver ", cn,3,3,1
    	w_action = objRSTxns.Recordcount

		'Response.Write "Count = : " & w_action & "<br>"
    	while w_action > 0

				iPlayerClaimed= objRSTxns.Fields("PID_Claimed").Value
				iPlayerWaived = objRSTxns.Fields("PID_Waived").Value
				iOwner        = objRSTxns.Fields("OwnerID").Value
				iPriority     = objRSTxns.Fields("WaiverPriority").Value
				iActivePlayers= objRSTxns.Fields("ActivePlayerCnt").Value

				if iPlayerWaived = 0 AND iActivePlayers >= 14 then
					'***************************************************************
					'REJECT THIS TRANSACTION BECAUSE THE PLAYER LIMIT IS 14 PER TEAM
					'****************************************************************
					TransType = "Waiver pick-up rejected (Roster full)"
					Cost = 0.00

					strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_
					iOwner & "', '" &  TransType & "', '" & iPlayerClaimed & "', '" &  Cost & "')"
					cn.Execute strSQL

					strSQL = "DELETE from tblWaivers where PID_Waived = 0 and PID_Claimed =" & iPlayerClaimed & " and OwnerID = " & iOwner & " ;"
					cn.Execute strSQL

				else
					'**********************************************************
					'UPDATE TO PLAYER TABLE for player being added.
					'**********************************************************
					strSQL = "update tblPlayers SET playerStatus = 'O', OwnerId = " & iOwner & ", " &_
									 "pendingwaiver = 0, clearwaiverdate = null, LastTeamInd = null " &_
									 "WHERE tblPlayers.PID = " & iPlayerClaimed & ";"
					cn.Execute strSQL

					'******************************************************************
					'UPDATE TO OWNERS TABLE.  Update other owners waiver priorities first
					'then set the current owner's waiver priority to 10.
					'******************************************************************
					strSQL ="update tblowners SET waiverpriority = waiverpriority - 1 WHERE waiverpriority > " & iPriority & ";"
					cn.Execute strSQL

					if iPlayerWaived = 0 then
						strSQL ="update tblowners SET waiverpriority = 10, ActivePlayerCnt = ActivePlayerCnt + 1 WHERE ownerid = " & iOwner & ";"
						cn.Execute strSQL
					else
						strSQL ="update tblowners SET waiverpriority = 10 WHERE ownerid = " & iOwner & ";"
						cn.Execute strSQL

						'**********************************************************
						'Update Player Table for Player being Waived
						'**********************************************************
						strSQL = "update tblPlayers " & _
										 "SET playerStatus = 'W', OwnerId = 0, pendingwaiver = 0, " & _
										 "LastTeamInd = " & iOwner & ", clearwaiverdate = date() + 1 " & _
										 "WHERE tblPlayers.PID = " & iPlayerWaived & ";"

						cn.Execute strSQL
						'**********************************************************
						'Player Released TRANSACTION
						'**********************************************************
						TransType = "Released"
						Cost = 0.00
						strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_
						iOwner & "', '" &  TransType & "', '" & iPlayerWaived & "', '" &  Cost & "')"
						cn.Execute strSQL
					end if

					'**********************************************************
					'Player Signed TRANSACTION
					'**********************************************************
					TransType = "Signed off waivers"
					Cost = 1.00
					strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_
					iOwner & "', '" &  TransType & "', '" &  iPlayerClaimed  & "', '" &  Cost & "')"
					cn.Execute strSQL

					objRejectWaivers.Open "SELECT * FROM tblWaivers where PID_Claimed =" & iPlayerClaimed & " and OwnerID <> " & iOwner & " ;" , cn
					TransType = "Waiver pick-up rejected"
					Cost = 0.00

					While Not objRejectWaivers.EOF

						iRejOwner = objRejectWaivers.Fields("OwnerID").Value
						iWaivedReject = objRejectWaivers.Fields("pid_waived").Value
						iClaimedReject = objRejectWaivers.Fields("pid_claimed").Value

						strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_
						iRejOwner & "', '" &  TransType & "', '" &  iClaimedReject & "', '" &  Cost & "')"
						cn.Execute strSQL

						'Update PendingWaiver flag
						strSQL = "update tblPlayers SET pendingwaiver = 0 WHERE tblPlayers.PID = " & iWaivedReject & ";"
						cn.Execute strSQL

						objRejectWaivers.MoveNext
					Wend

					objRejectWaivers.Close

					'*************************************************************************
					'Delete all entries from tblWaivers table where player_id = Player Claimed
					'*************************************************************************
					strSQL = "DELETE from tblWaivers where PID_Claimed = " & iPlayerClaimed & ";"
					cn.Execute strSQL

					'*************************************************************************
					'Delete any additional rows from the tblWaivers table for the player that
					'was just waived.  This is necessary if the owner had the same player on
					'multiple waivers.
					'*************************************************************************
					if iPlayerWaived <> 0 then
						strSQL = "DELETE from tblWaivers where PID_Waived = " & iPlayerWaived & ";"
						cn.Execute strSQL
					end if
			
				end if
			
						'**********************************************************
						'Close the Query and Open it again to see if any rows remain
						'**********************************************************
						ObjRsTxns.Close
						objRSTxns.Open		"SELECT * FROM qryUpdatewaiver ", cn,3,3,1
						w_action = objRSTxns.Recordcount

			wend

			ObjRsTxns.Close
		
		'*********************************************************************
		'Make players Free whose clearwaiver Date is less then Now()
		'and Set Rental Players back Free
		'*********************************************************************
		strSQL = "update tblPlayers " & _
		         "SET playerStatus = 'F', OwnerId = 0, clearwaiverdate = null, LastTeamInd = null " & _
		         "WHERE clearwaiverdate < now() and playerStatus = 'W'"

		cn.Execute strSQL

                '**************************************
                'CLEANUP when waiver stacking is used
                '**************************************
		strSQL = "update tblPlayers SET pendingwaiver = 0 "
 		cn.Execute strSQL

		'***************************************************************************
		'Set the time for the next pendingwaiversall run.  If tomorrow is a game
		'day, then set the pendingwaivers date to run 6 hours before cutofftime.
		'Note that that code subtracts 5 hours from the time because the times
		'in the database are CST but the server is hosted on EST.  If tomorrow is
		'not a game day then set the nextrun date to be 1:00 PM EST.
		'***************************************************************************
		objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = (date() + 1) ", cn,3,3,1

		if objNextRun.Recordcount > 0 then
			dnextrun = objNextRun.Fields("cutofftime").Value - 5/24
		else
			dnextrun = date() + 1 + (13/24)
		end if

		strSQL = 	"update tbltimedEvents " & _
							"SET lastrun = now(), nextrun = '"&dnextrun&"' " & _
							"WHERE event = 'pendingwaiversall' "

		cn.Execute strSQL

		objNextRun.Close
	end if

	objRSwaivers.Close	
 	
  '*******************************************************
	'Run setStaggeredAll event if it hasn't been run today.
	'*******************************************************
	objStagger.Open "SELECT * FROM tbltimedEvents " & _
	                  "where event = 'setStaggeredAll' and nextrun < now() ", cn,3,3,1
					 
	if  objStagger.Recordcount > 0 then	

	  strSQL = "update tblPlayers SET playerStatus = 'S' WHERE playerStatus = 'F' "
		cn.Execute strSQL
	
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
	       objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday >= (date() + 1) ", cn,3,3,1
  	    else
	       objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday >= date() ", cn,3,3,1
	    end if
		
        		
		if objNextRun.Recordcount > 0 then
			dnextrun = objNextRun.Fields("cutofftime").Value + 1/24
		else
			if wpos > 0 then
			 dnextrun = date() + 1 + (13/24)
			else
			 dnextrun = date() + (13/24)			
			end if
		end if
		
		strSQL = "update tbltimedEvents SET lastrun = now(), nextrun = '"&dnextrun&"' WHERE event = 'setStaggeredAll' "
		cn.Execute strSQL
		
		strSQL = "update tblParameterCtl set param_indicator = 1 where param_name = 'STAGGER_WINDOW' "
		cn.Execute strSQL
		
		objNextRun.Close
		
	end if  'End if objStagger.Recordcount > 0

	objStagger.Close
	
	'*****************************************************
	'Run setwaiversall event if it hasn't been run today.
	'*****************************************************
	objRSwaivers.Open "SELECT * FROM tbltimedEvents " & _
	                  "where event = 'setwaiversall' and nextrun < now() ", cn,3,3,1

	if  objRSwaivers.Recordcount > 0 then
	
    '********************************************************************
    'Check the time that this action is being triggered.  Either before or after
		'midnight.  If it is before midnight then next actions will be based on tomorrow.   
		'If after midnight then next actions will be based on today's date.		
		'If tomorrow is a game, then set the pendingwaivers date = to the cutofftimestagger.
		'Note that that code adds 1 hour to the time because the times
		'in the database are CST but the server is hosted on EST.  If tomorrow is
		'not a game day then set the nextrun date to be 7:00 PM EST.
		'*********************************************************************		
		
		wTime = time()
	    wpos = Instr(wTime, "PM")
	  
	    if wpos > 0 then   'Before Midnight
				dlineups = date()
				objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = (date() + 1) ", cn,3,3,1
				strSQL = "update tblPlayers " & _
		         "SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() + 1 " & _
		         "WHERE playerStatus = 'F' or playerStatus = 'S' or RentalPlayer = 1 "
			else
				dlineups = date() - 1
				objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = date() ", cn,3,3,1		
				strSQL = "update tblPlayers " & _
		         "SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() " & _
		         "WHERE playerStatus = 'F' or playerStatus = 'S' or RentalPlayer = 1 "
	    end if
				
		cn.Execute strSQL

		
		if objNextRun.Recordcount > 0 then
			dnextrun = objNextRun.Fields("cutofftimestagger").Value + 1/24
		else
			if wpos > 0 then
			 dnextrun = date() + 1 + (19/24)
			else
			 dnextrun = date() + (19/24)			
			end if
		end if

		strSQL = "update tbltimedEvents SET lastrun = now(), nextrun = '"&dnextrun&"' WHERE event = 'setwaiversall' "
		cn.Execute strSQL
		
		strSQL = "update tblParameterCtl set param_indicator = 0 where param_name = 'STAGGER_WINDOW' "
		cn.Execute strSQL		

		objNextRun.Close
		
		'********************************************************************************************
		'If today was a game day then set final_lineup flag and move lineups to tbl_lineups_staggered
		'********************************************************************************************
		objRSlineups.Open	"SELECT * FROM tbl_lineups where gameday = #"&dlineups&"# ", cn,3,3,1
		Response.write "Lineups today = "&objRSlineups.Recordcount&"<br>"
		if objRSlineups.Recordcount > 0 then
			strSQL = "update tbl_lineups set final_lineup = 1 where gameday = #"&dlineups&"# "
			cn.Execute strSQL
			
			strSQL = "delete from tbl_lineups_staggered"
			cn.Execute strSQL
			
			strSQL = "insert into tbl_lineups_staggered select * from tbl_lineups where gameday = #"&dlineups&"# "
			cn.Execute strSQL						
		end if
		
		objRSlineups.Close
	end if
	
	objRSwaivers.Close	
%>