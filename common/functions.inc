<%
	'***********************************************************
	' Remove_From_Lineups()
	' - Removes all players passed as parameters from future lineups.
	'
	'   <Called By>
	'   * tradeoffers.asp
	'   * transelect.asp
	'   * login.asp
	'   * maintainRoster.asp
	'***********************************************************
	Function Remove_From_Lineups (p1,p2,p3,p4,p5,p6,owner1,owner2)
	
	    Set RmlResults   = Server.CreateObject("ADODB.RecordSet")
		c_Time_1159      = "11:59:59 PM"
			
		RmlResults.Open "SELECT * FROM tbl_lineups where gameday >= date() AND ownerID in ("&owner1&","&owner2&") " & _
                                  "AND final_lineup = FALSE AND (" & _
					"sCenter in ("&p1&","&p2&","&p3&","&p4&","&p5&","&p6&")   OR " & _
					"sForward in ("&p1&","&p2&","&p3&","&p4&","&p5&","&p6&")  OR " & _
					"sForward2 in ("&p1&","&p2&","&p3&","&p4&","&p5&","&p6&") OR " & _
					"sGuard in ("&p1&","&p2&","&p3&","&p4&","&p5&","&p6&")    OR " & _
					"sGuard2 in ("&p1&","&p2&","&p3&","&p4&","&p5&","&p6&")  " & _
								       ") ",objConn,3,3,1
			
		'Response.Write "Inside Loop values "&p1&", "&p2&", "&p3&", "&p4&", "&p5&", "&p6&", "&owner1&", "&owner2&"<br>"
        'if RmlResults.RecordCount > 0 then
		'	Response.Write "Lineups to update = "&RmlResults.RecordCount&"<br>"
		'else
		'	Response.Write "*** NO LINEUPS TO UPDATE ***<br>"
        'end if		
				
		While Not RmlResults.EOF
		    
			'Response.Write "Inside Loop values p1,p2,p3,p4,p5,p6,owner1,owner2 "&p1&", "&p2&", "&p3&", "&p4&", "&p5&", "&p6&", "&owner1&", "&owner2&"<br>"
			'Response.Write "Inside Loop values p1,owner1 "&p1&", "&owner1&"<br>"
			
			wCenPID   = RmlResults.Fields("sCenter").Value
			wForPID   = RmlResults.Fields("sForward").Value
			wFor2PID  = RmlResults.Fields("sForward2").Value
			wGuaPID   = RmlResults.Fields("sGuard").Value
			wGua2PID  = RmlResults.Fields("sGuard2").Value
			wOwnerID  = RmlResults.Fields("OwnerID").Value
			wGameDay  = RmlResults.Fields("GameDay").Value
			
	       'Response.Write "Inside Loop values wCenPID,wForPID,wFor2PID,wGuaPID,wGua2PID,wOwnerID,wGameDay "&wCenPID&", "&wForPID&", "&wFor2PID&", "&wGuaPID&", "&wGua2PID&", "&wOwnerID&", "&wGameDay&"<br>"
	       'Response.Write "Inside Loop values wGuaPID, wGua2PID, wOwnerID, wGameDay "&wGuaPID&", "&wGua2PID&", "&wOwnerID&", "&wGameDay&"<br>"
					
			
		    if (CInt(wCenPID)=CInt(p1) OR CInt(wCenPID)=CInt(p2) OR CInt(wCenPID)=CInt(p3) OR CInt(wCenPID)=CInt(p4) OR CInt(wCenPID)=CInt(p5) OR CInt(wCenPID)=CInt(p6)) then
				strSQL = "update tbl_lineups set sCenter = 9998, sCenterBarps = 0, sCenterTip = '"&c_Time_1159&"' where ownerID = "&wOwnerID&" and gameday = #"&wGameDay&"#"
			    objConn.Execute strSQL		
				'Response.Write "strSQL = "&strSQL&"<br>"
			end if
			
			if (CInt(wForPID)=CInt(p1) OR CInt(wForPID)=CInt(p2) OR CInt(wForPID)=CInt(p3) OR CInt(wForPID)=CInt(p4) OR CInt(wForPID)=CInt(p5) OR CInt(wForPID)=CInt(p6)) then
				strSQL = "update tbl_lineups set sForward = 9996, sForwardBarps = 0, sForwardTip = '"&c_Time_1159&"' where ownerID = "&wOwnerID&" and gameday = #"&wGameDay&"#"
				objConn.Execute strSQL		
				'Response.Write "strSQL = "&strSQL&"<br>"				
			end if
			
			if (CInt(wFor2PID)=CInt(p1) OR CInt(wFor2PID)=CInt(p2) OR CInt(wFor2PID)=CInt(p3) OR CInt(wFor2PID)=CInt(p4) OR CInt(wFor2PID)=CInt(p5) OR CInt(wFor2PID)=CInt(p6)) then
				strSQL = "update tbl_lineups set sForward2 = 9997, sForward2Barps = 0, sForwardTip2 = '"&c_Time_1159&"' where ownerID = "&wOwnerID&" and gameday = #"&wGameDay&"#"
				objConn.Execute strSQL						
				'Response.Write "strSQL = "&strSQL&"<br>"
			end if
			
			if (CInt(wGuaPID)=CInt(p1) OR CInt(wGuaPID)=CInt(p2) OR CInt(wGuaPID)=CInt(p3) OR CInt(wGuaPID)=CInt(p4) OR CInt(wGuaPID)=CInt(p5) OR CInt(wGuaPID)=CInt(p6)) then
				strSQL = "update tbl_lineups set sGuard = 9994, sGuardBarps = 0, sGuardTip = '"&c_Time_1159&"' where ownerID = "&wOwnerID&" and gameday = #"&wGameDay&"#"				
				objConn.Execute strSQL		
				'Response.Write "strSQL = "&strSQL&"<br>"
			end if
			
			if (CInt(wGua2PID)=CInt(p1) OR CInt(wGua2PID)=CInt(p2) OR CInt(wGua2PID)=CInt(p3) OR CInt(wGua2PID)=CInt(p4) OR CInt(wGua2PID)=CInt(p5) OR CInt(wGua2PID)=CInt(p6)) then
				strSQL = "update tbl_lineups set sGuard2 = 9995, sGuard2Barps = 0, sGuardTip2 = '"&c_Time_1159&"' where ownerID = "&wOwnerID&" and gameday = #"&wGameDay&"#"
				objConn.Execute strSQL					
				'Response.Write "strSQL = "&strSQL&"<br>"				
			end if
							
		    RmlResults.MoveNext
		Wend
		
		RmlResults.Close
					
	End Function
	
	'*******************************************************************************************************
	' Reset_PendTrade_Flags()
	' - Reset All PendTrade Flags for all records to 0
	' - Next go thru the remaining pending trades and set the pendingTrade flag back to 1. 
	' - This was the quicked solution since it is hard to know how many other trades a player is involved in.
	'*******************************************************************************************************	
   	Function Reset_PendTrade_Flags()

	    Set objRSResetTradeFlag   = Server.CreateObject("ADODB.RecordSet")
		dim l_traded,l_traded2,l_traded3,l_acquired,l_acquired2,l_acquired3
		
		strSQL = "UPDATE tblPlayers SET PendingTrade = 0 "
		objConn.Execute strSQL

		objRSResetTradeFlag.Open "SELECT * FROM tblpendingtrades " , objConn

		While Not objRSResetTradeFlag.EOF

			l_traded   = objRSResetTradeFlag.Fields("tradedplayerid").Value
			l_traded2  = objRSResetTradeFlag.Fields("tradedplayerid2").Value
			l_traded3  = objRSResetTradeFlag.Fields("tradedplayerid3").Value
			l_acquired = objRSResetTradeFlag.Fields("acquiredplayerid").Value
			l_acquired2= objRSResetTradeFlag.Fields("acquiredplayerid2").Value
			l_acquired3= objRSResetTradeFlag.Fields("acquiredplayerid3").Value

			strSQL = "update tblPlayers SET pendingTrade = 1 " & _
			"WHERE PID in ("& l_traded &","& l_traded2 &","& l_traded3 &","& l_acquired &","& l_acquired2 &","& l_acquired3 &") "

			'Response.Write "strSQL = "&strSQL&" <br> "

			objConn.Execute strSQL
	        objRSResetTradeFlag.MoveNext
		Wend
		objRSResetTradeFlag.Close
	End Function

   '##################################
   ' Function get_top_barper
   '##################################
   Function get_top_barper(p_gameDate, p_ownerID, p_name, p_Maxbarps,pid)
   
      objRSWork.Open "Select * from newbox " & _
	                 "where gameDate = #"&p_gameDate&"# and (homeTeamoid = "&p_ownerID&" or awayTeamoid = "&p_ownerID&")", objConn,3,3,1
					 
      if objRSWork.RecordCount > 0 then
				if objRSWork.Fields("HomeTEamOID").Value = p_ownerID then
					if cdbl(objRSWork.Fields("hp1Barps").Value) > cdbl(objRSWork.Fields("hp2Barps").Value) then
						p_Maxbarps = objRSWork.Fields("hp1Barps").Value
						p_name     = left(objRSWork.Fields("hp1First").Value,1)&". "&left(objRSWork.Fields("hp1Last").Value,14)
						pid        = objRSWork.Fields("hp1PID").Value
					else
						p_Maxbarps = objRSWork.Fields("hp2Barps").Value
						p_name     = left(objRSWork.Fields("hp2First").Value,1)&". "&left(objRSWork.Fields("hp2Last").Value,14)
						pid        = objRSWork.Fields("hp2PID").Value
					end if
					if cdbl(objRSWork.Fields("hp3Barps").Value) > cdbl(p_Maxbarps) then
						p_Maxbarps = objRSWork.Fields("hp3Barps").Value
						p_name     = left(objRSWork.Fields("hp3First").Value,1)&". "&left(objRSWork.Fields("hp3Last").Value,14)
						pid        = objRSWork.Fields("hp3PID").Value
					end if
					if cdbl(objRSWork.Fields("hp4Barps").Value) > cdbl(p_Maxbarps) then
						p_Maxbarps = objRSWork.Fields("hp4Barps").Value
						p_name     = left(objRSWork.Fields("hp4First").Value,1)&". "&left(objRSWork.Fields("hp4Last").Value,14)
						pid        = objRSWork.Fields("hp4PID").Value
					end if
					if cdbl(objRSWork.Fields("hp5Barps").Value) > cdbl(p_Maxbarps) then
						p_Maxbarps = objRSWork.Fields("hp5Barps").Value
						p_name     = left(objRSWork.Fields("hp5First").Value,1)&". "&left(objRSWork.Fields("hp5Last").Value,14)
						pid        = objRSWork.Fields("hp5PID").Value
					end if		 		 
				else	  
	        if cdbl(objRSWork.Fields("ap1Barps").Value) > cdbl(objRSWork.Fields("ap2Barps").Value) then
						p_Maxbarps = objRSWork.Fields("ap1Barps").Value
						p_name     = left(objRSWork.Fields("ap1First").Value,1)&". "&left(objRSWork.Fields("ap1Last").Value,14)
					  pid        = objRSWork.Fields("ap1PID").Value
	        else
						p_Maxbarps = objRSWork.Fields("ap2Barps").Value
						p_name     = left(objRSWork.Fields("ap2First").Value,1)&". "&left(objRSWork.Fields("ap2Last").Value,14)
						pid        = objRSWork.Fields("ap2PID").Value
					end if
					if cdbl(objRSWork.Fields("ap3Barps").Value) > cdbl(p_Maxbarps) then
						p_Maxbarps = objRSWork.Fields("ap3Barps").Value
						p_name     = left(objRSWork.Fields("ap3First").Value,1)&". "&left(objRSWork.Fields("ap3Last").Value,14)
						pid        = objRSWork.Fields("ap3PID").Value
					end if
					if cdbl(objRSWork.Fields("ap4Barps").Value) > cdbl(p_Maxbarps) then
						p_Maxbarps = objRSWork.Fields("ap4Barps").Value
						p_name     = left(objRSWork.Fields("ap4First").Value,1)&". "&left(objRSWork.Fields("ap4Last").Value,14)
						pid        = objRSWork.Fields("ap4PID").Value
					end if
					if cdbl(objRSWork.Fields("ap5Barps").Value) > cdbl(p_Maxbarps) then
						p_Maxbarps = objRSWork.Fields("ap5Barps").Value
						p_name     = left(objRSWork.Fields("ap5First").Value,1)&". "&left(objRSWork.Fields("ap5Last").Value,14)
						pid        = objRSWork.Fields("ap5PID").Value
					end if
				end if	  
      else
			  p_name ("Box not Found")
			  p_barps = 0
			end if
   
      objRSWork.Close
   End Function	
	
%>	