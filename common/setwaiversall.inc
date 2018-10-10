<%

	Set SwaWaivers     = Server.CreateObject("ADODB.RecordSet")
	Set SwaLineups 	   = Server.CreateObject("ADODB.RecordSet")
	Set SwaNextRun	   = Server.CreateObject("ADODB.RecordSet")
	Set SwaNoGame	   = Server.CreateObject("ADODB.RecordSet")
	Set SwaStagger	   = Server.CreateObject("ADODB.RecordSet")
	Set objTxnAmt      = Server.CreateObject("ADODB.RecordSet")
	Set SsaEmail       = Server.CreateObject("ADODB.RecordSet")

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
		   SwaNextRun.Open  "SELECT * FROM qryGamedeadlines where gameday = (date() + 1) ", objConn,3,3,1
		   strSQL = "update tblPlayers " & _
		    		"SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() + 1 " & _
				    "WHERE playerStatus = 'F' or playerStatus = 'S' or RentalPlayer = 1 "
		else
		   dlineups = date() - 1
		   SwaNextRun.Open  "SELECT * FROM qryGamedeadlines where gameday = date() ", objConn,3,3,1		
		   strSQL = "update tblPlayers " & _
				    "SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() " & _
				    "WHERE playerStatus = 'F' or playerStatus = 'S' or RentalPlayer = 1 "
		end if
				
		objConn.Execute strSQL

		
		if SwaNextRun.Recordcount > 0 then
			dnextrun = SwaNextRun.Fields("cutofftimestagger").Value + 1/24
		else
			if wpos > 0 then
			   SwaNoGame.Open "SELECT min(TipTimeEst) as EarlyTip FROM tblLeagueSetup where gameDate = date() + 1 ", objConn
			   if ISNULL(SwaNoGame.Fields("EarlyTip").Value) then
			      dnextrun = date() + 1 + (19/24)
	           else
		          dnextrun = date() + 1 + SwaNoGame.Fields("EarlyTip").Value
	           end if 
			else
			   SwaNoGame.Open "SELECT min(TipTimeEst) as EarlyTip FROM tblLeagueSetup where gameDate = date() ", objConn
			   if ISNULL(SwaNoGame.Fields("EarlyTip").Value) then
			      dnextrun = date() + (19/24)	
	           else
		          dnextrun = date() + SwaNoGame.Fields("EarlyTip").Value
	           end if
			   		
			end if
			
			SwaNoGame.Close
		end if

		strSQL = "update tbltimedEvents SET lastrun_EST = now(), nextrun_EST = '"&dnextrun&"' WHERE event = 'setwaiversall' "
		objConn.Execute strSQL
		SwaNextRun.Close
		
		SwaStagger.Open "SELECT * From tblparameterCtl where param_indicator = TRUE and param_name = 'STAGGER_WINDOW' ", objConn,3,3,1
		
		if SwaStagger.Recordcount > 0 then		
			strSQL = "update tblParameterCtl set param_indicator = 0 where param_name = 'STAGGER_WINDOW' "
			objConn.Execute strSQL				   
	
		    wEmailOwnerID       = null
		    wAlert              = "receiveStaggerAlerts"
		    email_subject       = "Stagger Window Closed!" 
		    email_message       = "*** Players on your Roster are Open for Activity. IE: Trades & Waivers ***"

%>
		<!--#include virtual="Common/email_league.inc"-->
				   
<%

			'********************************************************************************
			'Beginning of LINEUP PENALTY LOGIC
			'********************************************************************************	

			'Dim w_action,iOwner

			Set swaOwners  = Server.CreateObject("ADODB.RecordSet")
			Set swaPlayers = Server.CreateObject("ADODB.RecordSet")
            Set swaWork    = Server.CreateObject("ADODB.RecordSet")
 
            
			swaWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PLAYOFF_START_DATE' ",objConn
	        wswaPlayoffStart = swaWork.Fields("param_date").Value
	        swaWork.Close

			swaOwners.Open	"SELECT * FROM tblowners " & _
											"WHERE ownerID <> 99 " & _
											"AND NOT EXISTS " & _
											" (select ownerid from tbl_lineups " & _
											"  where tbl_lineups.ownerId = tblowners.ownerID  AND  tbl_lineups.GameDay = #"&dlineups&"# ) " & _
											"order by ownerid", objConn,3,3,1	  
	 
			While Not swaOwners.EOF
				iOwner = swaOwners.Fields("OwnerID").Value  
	     
				'Response.Write "iOwner="&iOwner&" <br>" 
				G1 = 0 
				G1_BARPS = 0
				G1_POS = ""
				 
				G2 = 0 
				G2_BARPS = 0
				G2_POS = ""
				 
				F1 = 0 
				F1_BARPS = 0
				F1_POS = ""
				 
				F2 = 0 
				F2_BARPS = 0
				F2_POS = ""
		 
				CEN = 0		
				CEN_BARPS = 0
				CEN_POS = ""
			 
		        if (dlineups < wswaPlayoffStart) then
					swaPlayers.Open "SELECT * FROM qryMissingLineup " & _
													"WHERE gameday = #"&dlineups&"# " & _
													"AND ownerID = "&iOwner&" ", objConn,3,3,1
							
					While Not swaPlayers.EOF
				
						if swaPlayers.Fields("Pos").Value = "CEN" then
							if CEN = 0 then 
								CEN = swaPlayers.Fields("PID").Value
								CEN_BARPS = swaPlayers.Fields("barps").Value
								CEN_POS = swaPlayers.Fields("Pos").Value
							elseif CEN_POS = "F-C" then
								if F1 = 0 then   'Move F-C currently assigned to Center to the open Forward 1.  Assign this player to Center
									F1 = CEN
									F1_BARPS = CEN_BARPS
									F1_POS = CEN_POS						 
									CEN = swaPlayers.Fields("PID").Value
									CEN_BARPS = swaPlayers.Fields("barps").Value
									CEN_POS = swaPlayers.Fields("Pos").Value
								elseif F2 = 0 then  'Move F-C currently assigned to Center to the open Forward 2
									F2 = CEN
									F2_BARPS = CEN_BARPS
									F2_POS = CEN_POS						 
									CEN = swaPlayers.Fields("PID").Value
									CEN_BARPS = swaPlayers.Fields("barps").Value
									CEN_POS = swaPlayers.Fields("Pos").Value					  
								end if 
							end if
							 
						elseif swaPlayers.Fields("Pos").Value = "F-C" then
							if CEN = 0 then 
								CEN = swaPlayers.Fields("PID").Value
								CEN_BARPS = swaPlayers.Fields("barps").Value
								CEN_POS = swaPlayers.Fields("Pos").Value
							elseif F1 = 0 then   
								F1 = swaPlayers.Fields("PID").Value
								F1_BARPS = swaPlayers.Fields("barps").Value
								F1_POS = swaPlayers.Fields("Pos").Value
							elseif F2 = 0 then
								F2 = swaPlayers.Fields("PID").Value
								F2_BARPS = swaPlayers.Fields("barps").Value
								F2_POS = swaPlayers.Fields("Pos").Value
							end if				
							 
						elseif swaPlayers.Fields("Pos").Value = "FOR" then
							if F1 = 0 then 
								F1 = swaPlayers.Fields("PID").Value
								F1_BARPS = swaPlayers.Fields("barps").Value
								F1_POS = swaPlayers.Fields("Pos").Value
							elseif F2 = 0 then
								F2 = swaPlayers.Fields("PID").Value
								F2_BARPS = swaPlayers.Fields("barps").Value
								F2_POS = swaPlayers.Fields("Pos").Value		
							end if   
							 
						elseif swaPlayers.Fields("Pos").Value = "G-F" then   
							if G1 = 0 then
								G1 = swaPlayers.Fields("PID").Value
								G1_BARPS = swaPlayers.Fields("barps").Value
								G1_POS = swaPlayers.Fields("Pos").Value					  
							elseif G2 = 0 then
								G2 = swaPlayers.Fields("PID").Value
								G2_BARPS = swaPlayers.Fields("barps").Value
								G2_POS = swaPlayers.Fields("Pos").Value		
							elseif F1 = 0 then   
								F1 = swaPlayers.Fields("PID").Value
								F1_BARPS = swaPlayers.Fields("barps").Value
								F1_POS = swaPlayers.Fields("Pos").Value
							elseif F2 = 0 then
								F2 = swaPlayers.Fields("PID").Value
								F2_BARPS = swaPlayers.Fields("barps").Value
								F2_POS = swaPlayers.Fields("Pos").Value					  
							end if				
							 
						else
							if G1 = 0 then 
								G1 = swaPlayers.Fields("PID").Value
								G1_BARPS = swaPlayers.Fields("barps").Value
								G1_POS = swaPlayers.Fields("Pos").Value
							elseif G2 = 0 then
								G2 = swaPlayers.Fields("PID").Value
								G2_BARPS = swaPlayers.Fields("barps").Value
								G2_POS = swaPlayers.Fields("Pos").Value	
							elseif G2_POS = "G-F" then			
								if F1 = 0 then   'Move G-F currently assigned to Guard 2 to the open Forward 1.  Assign this player to Guard 2
									F1 = G2
									F1_BARPS = G2_BARPS
									F1_POS = G2_POS
									p_For1Name = p_Guard2Name
									G2 = swaPlayers.Fields("PID").Value
									G2_BARPS = swaPlayers.Fields("barps").Value
									G2_POS = swaPlayers.Fields("Pos").Value
									p_Guard2Name = swaPlayers.Fields("lastName").Value & " - " & G2_BARPS	
								elseif F2 = 0 then  'Move G-F currently assigned to Guard 2 to the open Forward 2.  Assign this player to Guard 2
									F2 = G2
									F2_BARPS = G2_BARPS
									F2_POS = G2_POS
									p_For2Name = p_Guard2Name
									G2 = swaPlayers.Fields("PID").Value
									G2_BARPS = swaPlayers.Fields("barps").Value
									G2_POS = swaPlayers.Fields("Pos").Value
									p_Guard2Name = swaPlayers.Fields("lastName").Value & " - " & G2_BARPS	
								end if										  
							elseif G1_POS = "G-F" then			
								if F1 = 0 then   'Move G-F currently assigned to Guard 1 to the open Forward 1.  Assign this player to Guard 1
									F1 = G1
									F1_BARPS = G1_BARPS
									F1_POS = G1_POS
									p_For1Name = p_Guard1Name
									G1 = swaPlayers.Fields("PID").Value
									G1_BARPS = swaPlayers.Fields("barps").Value
									G1_POS = swaPlayers.Fields("Pos").Value
									p_Guard1Name = swaPlayers.Fields("lastName").Value & " - " & G1_BARPS	
								elseif F2 = 0 then  'Move G-F currently assigned to Guard 1 to the open Forward 2.  Assign this player to Guard 1
									F2 = G1
									F2_BARPS = G1_BARPS
									F2_POS = G1_POS
									p_For2Name = p_Guard1Name
									G1 = swaPlayers.Fields("PID").Value
									G1_BARPS = swaPlayers.Fields("barps").Value
									G1_POS = swaPlayers.Fields("Pos").Value
									p_Guard1Name = swaPlayers.Fields("lastName").Value & " - " & G1_BARPS	
								end if					  
							end if   	   
						end if
						
						swaPlayers.MoveNext
					Wend
					swaPlayers.Close
				end if
					
				if G1 = 0 then 
					G1 = 9994
				end if
					 
				if G2 = 0 then
					G2 = 9994
				end if
					 
				if F1 = 0 then
					F1 = 9996
				end if
					 
				if F2 = 0 then
					F2 = 9996
				end if
					 
				if CEN = 0 then
					CEN = 9998
				end if

                strSQL = "insert into tbl_lineups(OwnerID,GameDay,sCenter,sCenterBarps,sForward,sForwardBarps,sForward2,sForward2Barps,sGuard,sGuardBarps,sGuard2,sGuard2Barps,penalty,sCenterTip,sForwardTip,sForwardTip2,sGuardTip,sGuardTip2) " &_
                         "values ('" &_
	    	              iOwner & "', '"&dlineups&"', '"&CEN&"', '"&CEN_BARPS&"','"&F1&"','"&F1_BARPS&"','"&F2&"','"&F2_BARPS&"','"&G1&"','"&G1_BARPS&"','"&G2&"','"&G2_BARPS&"',1,'11:59:59 PM','11:59:59 PM','11:59:59 PM','11:59:59 PM','11:59:59 PM')"
								 
				objConn.Execute strSQL
				
                strSQL = "insert into tbl_lineups_history(OwnerID,GameDay,sCenter,sCenterBarps,sForward,sForwardBarps,sForward2,sForward2Barps,sGuard,sGuardBarps,sGuard2,sGuard2Barps,penalty,sCenterTip,sForwardTip,sForwardTip2,sGuardTip,sGuardTip2) " &_
                         "values ('" &_
	    	              iOwner & "', '"&dlineups&"', '"&CEN&"', '"&CEN_BARPS&"','"&F1&"','"&F1_BARPS&"','"&F2&"','"&F2_BARPS&"','"&G1&"','"&G1_BARPS&"','"&G2&"','"&G2_BARPS&"',1,'11:59:59 PM','11:59:59 PM','11:59:59 PM','11:59:59 PM','11:59:59 PM')"
								 
				objConn.Execute strSQL
				
				'Response.Write "Sql = "&strSQL&" <br>"
			
				swaOwners.MoveNext
			Wend					
			swaOwners.Close
	
			
			'********************************************************************************
			'Ending of LINEUP PENALTY LOGIC
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

			objTxnAmt.Open "SELECT * FROM tblParameterCtl where param_name = 'RENT_PENALTY' ",objConn
			wCost = objTxnAmt.Fields("param_amount").Value
			objTxnAmt.Close
				
            strSQL = "update tblTransactions x set x.transType = 'Rental Not Played', x.transCost = "&wCost&" " & _
                     "where x.transtype = 'Rental' " & _
                     "and format(x.transdate, 'm/d/yyyy') = '"&dlineups&"' " & _
                     "and not exists " & _
                       "(select 1 " & _
                         "from tbl_lineups_staggered y " & _
                         "where x.ownerid = y.ownerid " & _
                         "and (x.transAddPlayer1 = y.sCenter or x.transAddPlayer1 = y.sForward or x.transAddPlayer1 = y.sForward2 or x.transAddPlayer1 = y.sGuard or x.transAddPlayer1 = y.sGuard2) " & _
                         "and gameday = #"&dlineups&"#)"
						 
			objConn.Execute strSQL

	
		end if
		
		SwaLineups.Close
	end if
	
	SwaWaivers.Close

%>