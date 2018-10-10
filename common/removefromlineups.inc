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
%>	