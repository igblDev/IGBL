<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%

On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn,objRS,ownerid,sAction,tradeid,objrstrade,objRSteam, objrsNames,txtNotes,towner, objRSOpponentName, objRSReg
	
	'EMAIL VARIABLES
	Dim email_to, email_subject, host, username, password, reply_to, port, from_address
	Dim first_name, last_name, home_address, email_from, telephone, comments, error_message
	Dim ObjSendMail, email_message, objEmail

	GetAnyParameter "Action", sAction
	GetAnyParameter "var_tradeid", stradeid
	txtNotes = Request.Form("txtNotes")

	Set objConn          = Server.CreateObject("ADODB.Connection")
	Set objRS            = Server.CreateObject("ADODB.RecordSet")
	Set objrstrade       = Server.CreateObject("ADODB.RecordSet")
	Set objrsNames       = Server.CreateObject("ADODB.RecordSet")
	Set objRSteam        = Server.CreateObject("ADODB.RecordSet")
	Set objRSOpponentName= Server.CreateObject("ADODB.RecordSet")
	Set objRSNext5       = Server.CreateObject("ADODB.RecordSet")
	Set objRSPO          = Server.CreateObject("ADODB.RecordSet")
	Set objRSReg         = Server.CreateObject("ADODB.RecordSet")
	Set objRSWork        = Server.CreateObject("ADODB.RecordSet")
	Set objEmail	     = Server.CreateObject("ADODB.RecordSet")
	
	'START FORECASTER LOGIC
	Dim OpponentName 
	Dim objRSHome, objRSAway, objRSAll,objRSPlayers
	Set objRSHome		= Server.CreateObject("ADODB.RecordSet")
	Set objRSAway		= Server.CreateObject("ADODB.RecordSet")
	Set objRSAll		= Server.CreateObject("ADODB.RecordSet")
	Set objRSPlayers    = Server.CreateObject("ADODB.RecordSet")	
	Set objRSReg   		= Server.CreateObject("ADODB.RecordSet")
	Set objRSPO    		= Server.CreateObject("ADODB.RecordSet")

	'END FORECASTER LOGIC

	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"

	%>
	<!--#include virtual="Common/session.inc"-->
	<%	
	tradeid   = stradeid
	objrstrade.Open	"SELECT t.*, to1.HomeEmail As TraderEmail, to1.TeamName AS TraderName, " & _
									"to2.HomeEmail, to2.TeamName, to2.ShortName " & _
									"FROM tblTradeAnalysis t, tblowners to1, tblowners to2 " & _
									"WHERE t.tradeid = " & tradeid & "  "  & _
									"and  t.tooid = to1.ownerid " & _
									"and  t.Fromoid = to2.ownerid ", objConn,3,3,1

	w_trade_count = objrstrade.Recordcount

	tplayer1  		= objrstrade.Fields("TradedPlayerID").Value
	tplayer2  		= objrstrade.Fields("TradedPlayerID2").Value
	tplayer3  		= objrstrade.Fields("TradedPlayerID3").Value
	aplayer1  		= objrstrade.Fields("AcquiredPlayerID").Value
	aplayer2  		= objrstrade.Fields("AcquiredPlayerID2").Value
	aplayer3  		= objrstrade.Fields("AcquiredPlayerID3").Value


	traderEmail 	= objrstrade.Fields("TraderEmail").Value
	traderName 		= objrstrade.Fields("TraderName").Value
	myname 				= objrstrade.Fields("ShortName").Value
	myemail 			= objrstrade.Fields("HomeEmail").Value
	towner 			  = objrstrade.Fields("toOID").Value
	tradeNotes    = Request.Form("txtNotes")
	tradepartner  = Request.Form("var_tradepartner")
	objrstrade.close

	select case sAction
	
		case "Submit Trade Offer"
	
		
		'Response.Write "TRADE NOTES = "&tradeNotes&" <br> "
		'Response.Write "Trade PT  = : " & tradepartner  & "  <br>" 
				objrstrade.Open	"SELECT * " & _
		        "FROM tblpendingtrades " & _
        		 "WHERE  FromOid = "& ownerid & "   " & _
        		     "and ToOid  = "& tradepartner & "   " & _
        		     "and tradedplayerid      = "& tplayer1  & "   " & _
                	 "and tradedplayerid2 = "& tplayer2  & "   " & _
	                 "and tradedplayerid3 = "& tplayer3  & "   " & _
	                 "and acquiredplayerid = "& aplayer1  & "  " & _
	                 "and acquiredplayerid2 = "& aplayer2 & " " & _
	                 "and acquiredplayerid3 = "& aplayer3  & " ", objConn,3,3,1


					'Response.Write "RECORD COUNT = "&objrstrade.RecordCount&" <br> "					
					if objrstrade.RecordCount > 0 then
						errorcode = "Invalid Offer"
					end if

					'Response.Write "error code = "&errorcode&" <br> "	
					objrstrade.close

					if errorcode = False then
			    	w_errorCt = 0

					objrstrade.Open	"SELECT * " & _
						"FROM tblplayers " & _
						"WHERE PID = "& tplayer1  & " ", objConn,3,3,1
					
					msgtradeplayers = objrstrade.Fields("firstName").Value & " " & objrstrade.Fields("lastName").Value & "<br />"   
					objrstrade.close

					objrstrade.Open	"SELECT * " & _
						"FROM tblplayers " & _
						"WHERE PID = "& aplayer1  & " ", objConn,3,3,1
					
					msgaquireplayers = objrstrade.Fields("firstName").Value & " " & objrstrade.Fields("lastName").Value & "<br />"  
					objrstrade.close
					
					if tplayer2 > 0 then
						objrstrade.Open	"SELECT * " & _
						"FROM tblplayers " & _
						"WHERE PID = "& tplayer2  & " ", objConn,3,3,1
						msgtradeplayers = msgtradeplayers + objrstrade.Fields("firstName").Value & " " & objrstrade.Fields("lastName").Value & "<br />"   
						objrstrade.close
					end if

					if tplayer3 > 0 then						
						objrstrade.Open	"SELECT * " & _
						"FROM tblplayers " & _
						"WHERE PID = "& tplayer3  & " ", objConn,3,3,1
						msgtradeplayers = msgtradeplayers + objrstrade.Fields("firstName").Value & " " & objrstrade.Fields("lastName").Value & "<br />"   
						objrstrade.close
					end if
					
					if aplayer2 > 0 then
						objrstrade.Open	"SELECT * " & _
						"FROM tblplayers " & _
						"WHERE PID = "& aplayer2  & " ", objConn,3,3,1					
						msgaquireplayers = msgaquireplayers + objrstrade.Fields("firstName").Value & " " & objrstrade.Fields("lastName").Value & "<br />"  
						objrstrade.close
					end if
					
					if aplayer3 > 0 then
						objrstrade.Open	"SELECT * " & _
						"FROM tblplayers " & _
						"WHERE PID = "& aplayer3  & " ", objConn,3,3,1					
						msgaquireplayers = msgaquireplayers + objrstrade.Fields("firstName").Value & " " & objrstrade.Fields("lastName").Value & "<br />"  
						objrstrade.close
					end if
					
					
					'** CHECK MY PLAYERS
					w_RetCd = Player_Owner_Relationship(tplayer1, ownerid, w_errorCt)
					w_RetCd = Player_Owner_Relationship(tplayer2, ownerid, w_errorCt)
					w_RetCd = Player_Owner_Relationship(tplayer3, ownerid, w_errorCt)

					'** CHECK TRADING PARTNER PLAYERS
					w_RetCd = Player_Owner_Relationship(aplayer1, tradepartner, w_errorCt)
					w_RetCd = Player_Owner_Relationship(aplayer2, tradepartner, w_errorCt)
					w_RetCd = Player_Owner_Relationship(aplayer3, tradepartner, w_errorCt)
					'	Response.Write "Error Count = "&w_errorCt&" <br> "


 					if w_errorCt > 0 then
	 					errorcode = "Invalid Offer"
					end if

					if errorcode = False then
						objrstrade.close
						errorcode = False
						wdays    = Request.Form("expireTime")
						strSQL ="insert into tblPendingTrades(FromOid,TradedPlayerID,TradedPlayerID2,TradedPlayerID3,ToOid,AcquiredPlayerID,AcquiredPlayerID2,AcquiredPlayerID3,DecisionDate) values ('" &_
						ownerid&"','"&tplayer1&"','"&tplayer2&"','"&tplayer3&"','"&tradepartner&"','"&aplayer1&"','"&aplayer2&"','"&aplayer3&"',now()-(1/24)+"&wdays&")"					
						objConn.Execute strSQL
						Response.Write "Insert into Pending Trade Table = "&strSQL&" <br> "
	  				'**********************************************************
						strSQL = "UPDATE tblPlayers SET tblPlayers.PendingTrade = Yes	WHERE PID=" & tplayer1 & " OR PID=" & tplayer2 & " OR PID=" & tplayer3 & "  OR  PID=" & aplayer1 & " or PID=" & aplayer2 & " or PID=" & aplayer3  &";"
						
						objConn.Execute strSQL
						
						strSQL = "DELETE FROM tblTradeAnalysis where tradeid =" & stradeid & " ;"
						objConn.Execute strSQL

						wEmailOwnerID       = tradepartner
						wAlert              = "receiveTradeAlerts"
						email_subject       = "Trade Offer from " & myname  
						email_message       = msgtradeplayers & "<br>"
						email_message       = email_message & "for <br><br>"
						email_message       = email_message & msgaquireplayers
						if len(tradeNotes) > 0 then
									email_message = email_message&"<br>"&tradeNotes&"<br>"
						end if 
						email_message       = email_message & "<br>www.igbl.org" 
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%								
					
					Response.Redirect "pendingtrades.asp?ownerid=" & ownerid

				end if

			end if

	  case "Forecast"
		conAction = "Continue"
		
		var_tradepartner = Request.Form("var_tradepartner")
		var_ownerid = Request.Form("var_ownerid")
		var_tradeid = Request.Form("var_tradeid")
		
   	case ""
		ownerid = session("ownerid")	

		case "Return"
		'ownerid = Request.querystring("txtOwnerID")
		ownerid = session("ownerid")	

		case "Counter"
		sURL = "tradeanalyzer.asp"
		conAction = "Revise"
		
		var_tradepartner = Request.Form("var_tradepartner")
		var_ownerid = Request.Form("var_ownerid")
		var_tradeid = Request.Form("var_tradeid")
		
		AddLinkParameter "var_ownerid", var_ownerid, sURL	
		AddLinkParameter "Action", conAction, sURL
		AddLinkParameter "cmbTeam", var_tradepartner, sURL
		AddLinkParameter "var_tradeid", var_tradeid, sURL
			
		Response.Redirect sURL

    case "Delete Invalid Trade"
		'Response.Write "Trade ID  = : " & stradeid  & "  <br>"
    strSQL = "DELETE FROM tblTradeAnalysis where tradeid =" & stradeid & " ;"
		objConn.Execute strSQL

		dim analysisCnt,objRSAnalysis
		
		Set objRSAnalysis  = Server.CreateObject("ADODB.RecordSet")	
		objRSAnalysis.Open  "Select * from tblTradeAnalysis where FromOid = "& ownerid & "", objConn,1,1
		analysisCnt   = objRSAnalysis.RecordCount
	  Response.Write "Record Count = : " & analysisCnt  & "  <br>"
		
		if analysisCnt > 0 then
			sURL = "pendingAnalyzedTrades.asp"
		else
			sURL = "dashboard.asp"
		end if	
		'sURL = "pendingAnalyzedTrades.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL

	end select
  '***********************************************************
	' Forecast_Lineup()
	' The Values of p_TraderID tells me which type of Query I should perform
	'    - If value is 0 then my result set should reflect my current roster and return how my lineups would look before the trade
	'    - If the value is not zero then my result set should reflect how my lineups would look after the trade.
	'***********************************************************	
	Function Forecast_Lineup (p_gameday, p_OwnerID, p_TraderID, p_CenName, CEN, CEN_BARPS, p_For1Name, F1, F1_BARPS, p_For2Name, F2, F2_BARPS, p_Guard1Name, G1, G1_BARPS, p_Guard2Name, G2, G2_BARPS)
                

		 p_Cen = ""
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
		 
				 
		 'Build optimum lineup from current roster
		 if p_TraderID = 0  then
		    objRSPlayers.Open "SELECT * FROM qryMissingLineup " &_
							   "WHERE gameday = #"&p_gameday&"# " &_
								 "AND ownerID = " & p_OwnerID & " " &_
								 "AND IR = 0 and Injury = 0 " , objConn,3,3,1		
		 'Build optimum based on proposed Trade					   
		 else  			
            objRSPlayers.Open "SELECT * FROM qryMissingLineup " &_
							   "WHERE gameday = #"&p_gameday&"# " &_
							   "AND ( " &_
							         "(ownerID = "& p_OwnerID & " AND PID not in (" & tplayer1 & "," & tplayer2 & "," & tplayer3 & ") AND IR = 0 and Injury = 0 ) " & _
							            "OR " &_									 
							         "(ownerID = "& p_TraderID & " AND PID in (" & aplayer1 & "," & aplayer2& "," & aplayer3 & ") ) " & _									 
                                "   ) " , objConn,3,3,1		 
		 end if		 'Build optimum lineup from current roster







		 
		 While Not objRSPlayers.EOF
			'Response.Write "Player = "&objRSPlayers.Fields("lastName").Value&", " & _
			'               "PID = "&objRSPlayers.Fields("PID").Value&", " & _
			'               "POS = "&objRSPlayers.Fields("Pos").Value&", " & _
			'               "BARPS = "&objRSPlayers.Fields("barps").Value&" <br> "
						   
			 if objRSPlayers.Fields("Pos").Value = "CEN" then
				   if CEN = 0 then 
					  CEN = objRSPlayers.Fields("PID").Value
					  CEN_BARPS = objRSPlayers.Fields("barps").Value
					  CEN_POS = objRSPlayers.Fields("Pos").Value
					  p_CenName = objRSPlayers.Fields("lastName").Value
				   elseif CEN_POS = "F-C" then
					  if F1 = 0 then   'Move F-C currently assigned to Center to the open Forward 1.  Assign this player to Center
						 F1 = CEN
						 F1_BARPS = CEN_BARPS
						 F1_POS = CEN_POS
						 p_For1Name = p_CenName				 
						 CEN = objRSPlayers.Fields("PID").Value
						 CEN_BARPS = objRSPlayers.Fields("barps").Value
						 CEN_POS = objRSPlayers.Fields("Pos").Value
						 p_CenName = objRSPlayers.Fields("lastName").Value
					  elseif F2 = 0 then  'Move F-C currently assigned to Center to the open Forward 2
						 F2 = CEN
						 F2_BARPS = CEN_BARPS
						 F2_POS = CEN_POS	
						 p_For2Name = p_CenName	
						 CEN = objRSPlayers.Fields("PID").Value
						 CEN_BARPS = objRSPlayers.Fields("barps").Value
						 CEN_POS = objRSPlayers.Fields("Pos").Value
						 p_CenName = objRSPlayers.Fields("lastName").Value						 
					  end if 
				   end if
				   
			 elseif objRSPlayers.Fields("Pos").Value = "F-C" then
				   if CEN = 0 then 
					  CEN = objRSPlayers.Fields("PID").Value
					  CEN_BARPS = objRSPlayers.Fields("barps").Value
					  CEN_POS = objRSPlayers.Fields("Pos").Value
					  p_CenName = objRSPlayers.Fields("lastName").Value
				   elseif F1 = 0 then   
					  F1 = objRSPlayers.Fields("PID").Value
					  F1_BARPS = objRSPlayers.Fields("barps").Value
					  F1_POS = objRSPlayers.Fields("Pos").Value
					  p_For1Name = objRSPlayers.Fields("lastName").Value
				   elseif F2 = 0 then
					  F2 = objRSPlayers.Fields("PID").Value
					  F2_BARPS = objRSPlayers.Fields("barps").Value
					  F2_POS = objRSPlayers.Fields("Pos").Value
					  p_For2Name = objRSPlayers.Fields("lastName").Value
				   end if				
				   
			 elseif objRSPlayers.Fields("Pos").Value = "FOR" then
				   if F1 = 0 then 
					  F1 = objRSPlayers.Fields("PID").Value
					  F1_BARPS = objRSPlayers.Fields("barps").Value
					  F1_POS = objRSPlayers.Fields("Pos").Value
					  p_For1Name = objRSPlayers.Fields("lastName").Value
				   elseif F2 = 0 then
					  F2 = objRSPlayers.Fields("PID").Value
					  F2_BARPS = objRSPlayers.Fields("barps").Value
					  F2_POS = objRSPlayers.Fields("Pos").Value	
					  p_For2Name = objRSPlayers.Fields("lastName").Value
				   end if   
				   
			elseif objRSPlayers.Fields("Pos").Value = "G-F" then
				   if G1 = 0 then
					  G1 = objRSPlayers.Fields("PID").Value
					  G1_BARPS = objRSPlayers.Fields("barps").Value
					  G1_POS = objRSPlayers.Fields("Pos").Value	
					  p_Guard1Name = objRSPlayers.Fields("lastName").Value
				   elseif G2 = 0 then
					  G2 = objRSPlayers.Fields("PID").Value
					  G2_BARPS = objRSPlayers.Fields("barps").Value
					  G2_POS = objRSPlayers.Fields("Pos").Value	
					  p_Guard2Name = objRSPlayers.Fields("lastName").Value
				   elseif F1 = 0 then   
					  F1 = objRSPlayers.Fields("PID").Value
					  F1_BARPS = objRSPlayers.Fields("barps").Value
					  F1_POS = objRSPlayers.Fields("Pos").Value
					  p_For1Name = objRSPlayers.Fields("lastName").Value
				   elseif F2 = 0 then
					  F2 = objRSPlayers.Fields("PID").Value
					  F2_BARPS = objRSPlayers.Fields("barps").Value
					  F2_POS = objRSPlayers.Fields("Pos").Value
					  p_For2Name = objRSPlayers.Fields("lastName").Value
				   end if												  
			else  'Guard Logic
				   if G1 = 0 then 
					  G1 = objRSPlayers.Fields("PID").Value
					  G1_BARPS = objRSPlayers.Fields("barps").Value
					  G1_POS = objRSPlayers.Fields("Pos").Value
					  p_Guard1Name = objRSPlayers.Fields("lastName").Value
				   elseif G2 = 0 then
					  G2 = objRSPlayers.Fields("PID").Value
					  G2_BARPS = objRSPlayers.Fields("barps").Value
					  G2_POS = objRSPlayers.Fields("Pos").Value
 					      p_Guard2Name = objRSPlayers.Fields("lastName").Value
				   elseif G2_POS = "G-F" then			
					  if F1 = 0 then   'Move G-F currently assigned to Guard 2 to the open Forward 1.  Assign this player to Guard 2
						 F1 = G2
						 F1_BARPS = G2_BARPS
						 F1_POS = G2_POS
						 p_For1Name = p_Guard2Name
						 G2 = objRSPlayers.Fields("PID").Value
						 G2_BARPS = objRSPlayers.Fields("barps").Value
						 G2_POS = objRSPlayers.Fields("Pos").Value
						 p_Guard2Name = objRSPlayers.Fields("lastName").Value
					  elseif F2 = 0 then  'Move G-F currently assigned to Guard 2 to the open Forward 2.  Assign this player to Guard 2
						 F2 = G2
						 F2_BARPS = G2_BARPS
						 F2_POS = G2_POS
						 p_For2Name = p_Guard2Name
						 G2 = objRSPlayers.Fields("PID").Value
						 G2_BARPS = objRSPlayers.Fields("barps").Value
						 G2_POS = objRSPlayers.Fields("Pos").Value
						 p_Guard2Name = objRSPlayers.Fields("lastName").Value
					  end if										  
				   elseif G1_POS = "G-F" then			
					  if F1 = 0 then   'Move G-F currently assigned to Guard 1 to the open Forward 1.  Assign this player to Guard 1
						 F1 = G1
						 F1_BARPS = G1_BARPS
						 F1_POS = G1_POS
						 p_For1Name = p_Guard1Name
						 G1 = objRSPlayers.Fields("PID").Value
						 G1_BARPS = objRSPlayers.Fields("barps").Value
						 G1_POS = objRSPlayers.Fields("Pos").Value
						 p_Guard1Name = objRSPlayers.Fields("lastName").Value
					  elseif F2 = 0 then  'Move G-F currently assigned to Guard 1 to the open Forward 2.  Assign this player to Guard 1
						 F2 = G1
						 F2_BARPS = G1_BARPS
						 F2_POS = G1_POS
						 p_For2Name = p_Guard1Name
						 G1 = objRSPlayers.Fields("PID").Value
						 G1_BARPS = objRSPlayers.Fields("barps").Value
						 G1_POS = objRSPlayers.Fields("Pos").Value
						 p_Guard1Name = objRSPlayers.Fields("lastName").Value
					  end if												  
				   end if   	   
			 end if			
						
			objRSPlayers.MoveNext
		 Wend
		 objRSPlayers.Close
		 
		 if IsNull(CEN_BARPS)      then CEN_BARPS   = 0 end if
	     if IsNull(F1_BARPS)       then F1_BARPS    = 0 end if
	     if IsNull(F2_BARPS)       then F2_BARPS    = 0 end if
	     if IsNull(G1_BARPS)       then G1_BARPS    = 0 end if
	     if IsNull(G2_BARPS)       then G2_BARPS    = 0 end if
		 
		 if F2_BARPS > F1_BARPS then
			w_Name       = p_For1Name
			p_For1Name   = p_For2Name
			p_For2Name   = w_Name
						
			w_PID        = F1
			F1           = F2
			F2           = w_PID			
			
			w_Barps      = F1_BARPS
			F1_BARPS     = F2_BARPS
			F2_BARPS     = w_Barps
		 end if
		 
		 if G2_BARPS > G1_BARPS then
			w_Name       = p_Guard1Name
			p_Guard1Name = p_Guard2Name
			p_Guard2Name = w_Name
			
			w_PID        = G1
			G1           = G2
			G2           = w_PID

	        w_Barps      = G1_BARPS
			G1_BARPS     = G2_BARPS
			G2_BARPS     = w_Barps			
		 end if
		 
	End Function

	'***********************************************************
	' Get_Player_Names()
	'***********************************************************
	Function Get_Player_Names(p_tradeID, msgtradeplayers, msgacquireplayers)

		objrsNames.Open "SELECT * FROM qryAnalysisTrade Where TradeID = "&p_tradeID&" " , objConn

		'********************************************************
		'** Populate Traded Players String
		'********************************************************
		msgtradeplayers = objrsNames.Fields("t1first").Value & " " & objrsNames.Fields("t1last").Value & "<br />"

		if objrsNames.Fields("t2PID").Value > 0 then
			msgtradeplayers = msgtradeplayers & objrsNames.Fields("t2first").Value & " " & objrsNames.Fields("t2last").Value & "<br />"  
		end if

		if objrsNames.Fields("t3PID").Value > 0 then
			msgtradeplayers = msgtradeplayers & objrsNames.Fields("t3first").Value & " " & objrsNames.Fields("t3last").Value & "<br />"
		end if

		'********************************************************
		'** Populate Acquired Players String
		'********************************************************

		msgacquireplayers= objrsNames.Fields("a1first").Value & " " & objrsNames.Fields("a1last").Value & "<br />"

		if objrsNames.Fields("a2PID").Value > 0 then
			msgacquireplayers = msgacquireplayers & objrsNames.Fields("a2first").Value & " " & objrsNames.Fields("a2last").Value & "<br />"
		end if

		if objrsNames.Fields("a3PID").Value > 0 then
			msgacquireplayers = msgacquireplayers &  objrsNames.Fields("a3first").Value & " " & objrsNames.Fields("a3last").Value & "<br />"
		end if

		objrsNames.Close
	End Function
	
	'***********************************************************
	' Get_One_Name()
	'***********************************************************
   	Function Get_One_Name(playerID, playerName, lName, playerGmCnt, playerPOGMCnt)
	    'Response.Write "Inside Function. playerID = "&playerID
		objrsNames.Open "SELECT * FROM tblPlayers Where PID = "&playerID&" " , objConn
		playerName = left(objrsNames.Fields("firstName").Value,1) & ". " & left(objrsNames.Fields("lastName").Value,14)
		lName = objrsNames.Fields("lastName").Value
        objrsNames.Close
		
		objRSReg.Open  		"SELECT * FROM qryAllPlayerGameDays WHERE pid = " & playerID & " and gameday >= Date() " & _
							"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') order by gameday ", objConn,1,1					 

		playerGmCnt = objRSReg.RecordCount
		objRSReg.Close
		
		objRSPO.Open 		"select gameday from qryMissingLineup " & _
							"where pid = " & playerID & " " & _
							"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1  												
		
		playerPOGMCnt = objRSPO.RecordCount							
		objRSPO.Close
		
		'Response.Write ", playerName = "&playerName&"<br>"
	End Function

	objRS.Open "SELECT * FROM qryAnalysisTrade WHERE (((qryAnalysisTrade.towner)=" & ownerid & "))", objConn,3,3,1

	t2pid = 0
	a2pid = 0
	t3pid = 0
	a3pid = 0	
	
	t1first  = objRS.Fields("t1first").Value
	t1last   = objRS.Fields("t1last").Value
	t1pid    = objRS.Fields("t1PID").Value

	if objRS.Fields("t2PID").Value > 0 then
		t2first= objRS.Fields("t2first").Value
		t2last = objRS.Fields("t2last").Value
	end if

	if objRS.Fields("t3PID").Value > 0 then
		t3first= objRS.Fields("t3first").Value
		t3last = objRS.Fields("t3last").Value
	end if

	a1first  = objRS.Fields("a1first").Value
	a1last   = objRS.Fields("a1last").Value
	a1pid    = objRS.Fields("a1PID").Value

	if objRS.Fields("a2PID").Value > 0 then
		a2first= objRS.Fields("a2first").Value
		a2last = objRS.Fields("a2last").Value
	end if

	if objRS.Fields("a3PID").Value > 0 then
		a3first= objRS.Fields("a3first").Value
		a3last = objRS.Fields("a3last").Value
	end if

	

%>
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
white {
	color: white;
}

black {
	color:black;
}

green {
	color:#468847;
	font-weight: 500;
}

td {
	vertical-align:middle;

}
th {
	vertical-align:middle;

}
.alert-analysis {
    font-weight: bold;
    color: black;
    background-color: #dddddd;
    border-color: black;
    border-radius: 24px;
		border-width: medium;
}
table.box{
	border-bottom-color:black;
	border-bottom-style:double;
	border-bottom-width:thick;
	border-left-color:black;
	border-left-style:solid;
	border-left-width:thin;
	border-right-color:black;
	border-right-style:solid;
	border-right-width:thin;
}
.h5, .h6, h5, h6 {
    font-family: inherit;
    font-weight: 500;
    line-height: 1.1;
    color: inherit;
    text-align: center;
}
.h4, h4 {
    font-size: 16px;
    color: #212121;
    font-weight: 600;
		text-transform: uppercase;
}

redText {
	color:#9a1400;	
	font-weight:500;
}

</style>
</head>
<body>
<script language="JavaScript" type="text/javascript">
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip(); 
});	
function processReforecast(theForm) {
		
	var reForecastCnt = null;
	reForecastCnt = theForm.elements["reForcast"].value;

	if (reForecastCnt == "") {
		alert("Select Neutral Factor Indicator" ); 
		return false;
	}
	return (true);
}	
</script>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Pending Analysed Trades</strong>
			</div>
		</div>
	</div>
</div>
<% if sAction = "" then %>
<%
 While Not objRS.EOF
	wTeamName = replace((objRS.Fields("ateamname").Value), "THE ", "")
	if len(wTeamName) >19 then 
	 wTeamName = objRS.Fields("tteamnameshort").Value
	end if
%>
<form action="pendingAnalyzedTrades.asp" name="frmMain" method="POST">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_tradeid" value="<%=objRS.Fields("tradeid").Value %>" />
<input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("aowner").Value %>" />
<input type="hidden" name="var_invtradeind" value="<%=objRS.Fields("InvalidTradeInd").Value %>" />
<div class="container">
<div class="row">
	<div class="col-md-12 col-sm-12 col-xs-12">	
			<div class="panel panel-override">
				<table class="table table-custom-black  table-bordered table-condensed">
					<tr>
						<th style="width:50%;text-align:left;">My Players</th>
						<th style="width:50%;text-align:left;"	><%=wTeamName%></th>
					</tr>
							<tr>	
								<td style="background-color:white;"> 
									<table class="table table-bordered table-condensed">
										<%
										objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t1first&"' and lastName = '"&t1last&"' "
										t1playerBarps = 0	 
										t1blks   = 0
										t1asts   = 0
										t1rebs   = 0
										t1pts    = 0 
										t1stls   = 0
										t1turns  = 0
										t1threes = 0
										if objrstrade.RecordCount > 0 then
											t1playerBarps = objrstrade.Fields("barps").Value
											t1blks      	= objrstrade.Fields("blk").Value
											t1asts      	= objrstrade.Fields("ast").Value
											t1rebs      	= objrstrade.Fields("reb").Value
											t1pts       	= objrstrade.Fields("ppg").Value 
											t1stls      	= objrstrade.Fields("stl").Value
											t1turns     	= objrstrade.Fields("to").Value
											t1threes    	= objrstrade.Fields("three").Value

											if t1pid > 0 then
												objRSNext5.Open "select gameday from qryMissingLineup " & _
																				"where pid = " & t1pid & " " & _
																				"and gameday >= Date() " & _
																				"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
																 
												objRSPO.Open 	"select gameday from qryMissingLineup " & _
																			"where pid = " & t1pid & " " & _
																			"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
															 
												tplayer1GmCnt  = objRSNext5.RecordCount
												tplayer1POGmCnt= objRSPO.RecordCount
												objRSNext5.close
												objRSPO.close
											end if					
										end if
										%>
													<tr style="background-color:white;"> 
														<td>
															<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t1last").Value,10) %></a>
															<span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;<span ><%=round(t1playerBarps,2) %></span>
														</td>
													</tr>
										<%
										objrstrade.Close
										t2blks   = 0
										t2asts   = 0
										t2rebs   = 0
										t2pts    = 0 
										t2stls   = 0
										t2turns  = 0
										t2threes = 0
										t2playerBarps = 0	
										%>
										<%	
										if objRS.Fields("t2PID").Value > 0 then
											objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t2first&"' and lastName = '"&t2last&"' "
											if objrstrade.RecordCount > 0 then
												t2playerBarps = objrstrade.Fields("barps").Value
												t2blks      = objrstrade.Fields("blk").Value
												t2asts      = objrstrade.Fields("ast").Value
												t2rebs      = objrstrade.Fields("reb").Value
												t2pts       = objrstrade.Fields("ppg").Value 
												t2stls      = objrstrade.Fields("stl").Value
												t2turns     = objrstrade.Fields("to").Value
												t2threes    = objrstrade.Fields("three").Value
												t2pid       = objRS.Fields("t2PID").Value
												
												if t2pid > 0 then
														objRSNext5.Open "select gameday from qryMissingLineup " & _
																						"where pid = " & t2pid & " " & _
																						"and gameday >= Date() " & _
																						"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
																 
															objRSPO.Open 	"select gameday from qryMissingLineup " & _
																						"where pid = " & t2pid & " " & _
																						"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
														
													tplayer2GmCnt  = objRSNext5.RecordCount
													tplayer2POGmCnt= objRSPO.RecordCount
													objRSNext5.close
													objRSPO.close
												end if														
											end if
											%>
											<tr style="background-color:white;"> 
												<td>
													<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t2last").Value,10) %></a>
													<span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;<span ><%=round(t2playerBarps,2) %></span>
												</td>
											</tr>	
											<%end if%>
										<%
										objrstrade.Close
										t3blks   = 0
										t3asts   = 0
										t3rebs   = 0
										t3pts    = 0 
										t3stls   = 0
										t3turns  = 0
										t3threes = 0	
										t3playerBarps = 0	
										if (objRS.Fields("t3PID").Value) > 0 then
											objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t3first&"' and lastName = '"&t3last&"' "
											if objrstrade.RecordCount > 0 then
												t3playerBarps = objrstrade.Fields("barps").Value
												t3blks      = objrstrade.Fields("blk").Value
												t3asts      = objrstrade.Fields("ast").Value
												t3rebs      = objrstrade.Fields("reb").Value
												t3pts       = objrstrade.Fields("ppg").Value 
												t3stls      = objrstrade.Fields("stl").Value
												t3turns     = objrstrade.Fields("to").Value
												t3threes    = objrstrade.Fields("three").Value
												t3pid       = objRS.Fields("t3PID").Value
												
												objRSNext5.Open  "select gameday from qryMissingLineup " & _
																				 "where pid = " & t3pid & " " & _
																				 "and gameday >= Date() " & _
																				 "and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
																 
													objRSPO.Open 	"select gameday from qryMissingLineup " & _
																				"where pid = " & t3pid & " " & _
																				"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1  
													

												tplayer3GmCnt  = objRSNext5.RecordCount
												tplayer3POGmCnt= objRSPO.RecordCount
												objRSNext5.close
												objRSPO.close
																					
											end if
										%>
										<tr style="background-color:white;"> 
											<td>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t3last").Value,10) %></a>
												<span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;<span ><%=round(t3playerBarps,2) %></span>
											</td>
										</tr>	
										<% end if %>								
									</table>
								</td>
								
								<td style="background-color:white;"> 
									<table class="table table-bordered table-condensed">
						
					<%
					objrstrade.Close
					a1blks   = 0
					a1asts   = 0
					a1rebs   = 0
					a1pts    = 0 
					a1stls   = 0
					a1turns  = 0
					a1threes = 0
					a1playerBarps = 0	 
					objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a1first&"' and lastName = '"&a1last&"' " 
					if objrstrade.RecordCount > 0 then
						a1playerBarps = objrstrade.Fields("barps").Value
						a1blks      = objrstrade.Fields("blk").Value
						a1asts      = objrstrade.Fields("ast").Value
						a1rebs      = objrstrade.Fields("reb").Value
						a1pts       = objrstrade.Fields("ppg").Value 
						a1stls      = objrstrade.Fields("stl").Value
						a1turns     = objrstrade.Fields("to").Value
						a1threes    = objrstrade.Fields("three").Value		
																			
						objRSNext5.Open "select gameday from qryMissingLineup " & _
														"where pid = " & a1pid & " " & _
														"and gameday >= Date() " & _
														"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
										 
						objRSPO.Open 		"select gameday from qryMissingLineup " & _
														"where pid = " & a1pid & " " & _
														"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1  												
									
					aplayer1GmCnt  = objRSNext5.RecordCount
					aplayer1POGmCnt= objRSPO.RecordCount																										
					objRSNext5.close
					objRSPO.close
				end if
				%>
					<tr  style="background-color:white">
						<td>
							<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a1last").Value,10) %></a>
							<span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;<span ><%=round(a1playerBarps,2) %></span>
						</td>
					</tr>	 

					<%
					objrstrade.Close
					a2blks   = 0
					a2asts   = 0
					a2rebs   = 0
					a2pts    = 0 
					a2stls   = 0
					a2turns  = 0
					a2threes = 0
					a2playerBarps = 0	 
					if (objRS.Fields("a2PID").Value) > 0 then
						objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a2first&"' and lastName = '"&a2last&"' " 
						if objrstrade.RecordCount > 0 then
							a2playerBarps = objrstrade.Fields("barps").Value
							a2blks      = objrstrade.Fields("blk").Value
							a2asts      = objrstrade.Fields("ast").Value
							a2rebs      = objrstrade.Fields("reb").Value
							a2pts       = objrstrade.Fields("ppg").Value 
							a2stls      = objrstrade.Fields("stl").Value
							a2turns     = objrstrade.Fields("to").Value
							a2threes    = objrstrade.Fields("three").Value
							a2pid       = objRS.Fields("a2PID").Value
											 
							objRSNext5.Open  	"select gameday from qryMissingLineup " & _
																"where pid = " & a2pid & " " & _
																"and gameday >= Date() " & _
																"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
											 
							objRSPO.Open 			"select gameday from qryMissingLineup " & _
																"where pid = " & a2pid & " " & _
																"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1  
										 
							aplayer2GmCnt  = objRSNext5.RecordCount
							aplayer2POGmCnt= objRSPO.RecordCount																										
							objRSNext5.close
							objRSPO.close									
						end if
					%>
										<tr  style="background-color:white">
											<td>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a2last").Value,10) %></a>
												<span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;<span ><%=round(a2playerBarps,2) %></span>
											</td>
										</tr>	 

		
						<% end if %>		
					<% 
					objrstrade.Close
					a3blks   = 0
					a3asts   = 0
					a3rebs   = 0
					a3pts    = 0 
					a3stls   = 0
					a3turns  = 0
					a3threes = 0	
					a3playerBarps = 0	
					if (objRS.Fields("a3PID").Value) > 0 then
					objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a3first&"' and lastName = '"&a3last&"' "  		
					if objrstrade.RecordCount > 0 then
						a3playerBarps = objrstrade.Fields("barps").Value
						a3blks        = objrstrade.Fields("blk").Value
						a3asts        = objrstrade.Fields("ast").Value
						a3rebs        = objrstrade.Fields("reb").Value
						a3pts         = objrstrade.Fields("ppg").Value 
						a3stls        = objrstrade.Fields("stl").Value
						a3turns       = objrstrade.Fields("to").Value
						a3threes      = objrstrade.Fields("three").Value	
						a3pid         = objRS.Fields("a3PID").Value
															
						objRSNext5.Open "select gameday from qryMissingLineup " & _
														"where pid = " & a3pid & " " & _
														"and gameday >= Date() " & _
														"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
											 
						objRSPO.Open 		"select gameday from qryMissingLineup " & _
														"where pid = " & a3pid & " " & _
														"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1  												 
																						 
						aplayer3GmCnt  = objRSNext5.RecordCount
						aplayer3POGmCnt= objRSPO.RecordCount																										
						objRSNext5.close
						objRSPO.close	
					end if
					%>
 
					<tr  style="background-color:white">
						<td>
							<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a3last").Value,10) %></a>
							<span class="greenTrade big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;<span ><%=round(a3playerBarps,2) %></span>
						</td>
					</tr>
						<% end if %>								
									</table>
								</td>
							</tr>



					<%
					objrstrade.Close 
					
						tbarps      = cDbl(t1playerBarps) + cDbl(t2playerBarps) + cDbl(t3playerBarps)
						abarps      = cDbl(a1playerBarps) + cDbl(a2playerBarps) + cDbl(a3playerBarps)							
						diffBarps   = cDbl(abarps) -  cDbl(tbarps)
						
						tblks       = cDbl(t1blks) + cDbl(t2blks) + cDbl(t3blks)
						ablks       = cDbl(a1blks) + cDbl(a2blks) + cDbl(a3blks)
						diffblks    = cDbl(ablks) - cDbl(tblks)
						
						tasts       = cDbl(t1asts) + cDbl(t2asts) + cDbl(t3asts)
						aasts       = cDbl(a1asts) + cDbl(a2asts) + cDbl(a3asts)
						diffasts    = cDbl(aasts) - cDbl(tasts)

						trebs       = cDbl(t1rebs) + cDbl(t2rebs) + cDbl(t3rebs)
						arebs       = cDbl(a1rebs) + cDbl(a2rebs) + cDbl(a3rebs)
						diffrebs    = cDbl(arebs) - cDbl(trebs)								

						tpts        = cDbl(t1pts) + cDbl(t2pts) + cDbl(t3pts)
						apts        = cDbl(a1pts) + cDbl(a2pts) + cDbl(a3pts)
						diffpts     = cDbl(apts) - cDbl(tpts)

						tstls       = cDbl(t1stls) + cDbl(t2stls) + cDbl(t3stls)
						astls       = cDbl(a1stls) + cDbl(a2stls) + cDbl(a3stls)
						diffstls    = cDbl(astls) - cDbl(tstls)

						tstls       = cDbl(t1stls) + cDbl(t2stls) + cDbl(t3stls)
						astls       = cDbl(a1stls) + cDbl(a2stls) + cDbl(a3stls)
						diffstls    = cDbl(astls) - cDbl(tstls)

						ttos        = cDbl(t1turns) + cDbl(t2turns) + cDbl(t3turns)
						atos        = cDbl(a1turns) + cDbl(a2turns) + cDbl(a3turns)
						difftos     = cDbl(ttos) - cDbl(atos)

						tthrees     = cDbl(t1threes) + cDbl(t2threes) + cDbl(t3threes)
						athrees     = cDbl(a1threes) + cDbl(a2threes) + cDbl(a3threes)
						diffthrees  = cDbl(athrees) - cDbl(tthrees)
					%>
					</table>
					<table class="table table-custom-black  table-bordered table-condensed">
					<tr style="text-align: center;">
						<td class="big" style="background-color:#212121;font-size:14px;color:#FFEB3B;font-weight:700;" colspan="8">Plus/Minus By Stats</td>
					</tr>
					<tr>
						<th class="big text-center" style="width:14%">B/pg</th>
						<th class="big text-center" style="width:14%">A/pg</th>
						<th class="big text-center" style="width:14%">R/pg</th>
						<th class="big text-center" style="width:14%">P/pg</th>
						<th class="big text-center" style="width:14%">S/pg</th>
						<th class="big text-center" style="width:14%">3/pg</th>
						<th class="big text-center" style="width:16%">T/pg</th>	
						</tr>	
						<tr style="background-color:#212121;font-size:14px;color:#FFEB3B;font-weight:700;text-align: center;">
						<% if (cDbl(tblks)) < (cDbl(ablks)) then %>
								<td  class="big" align="center">+<%=round(diffblks,1)%></td>
						<% elseif (cDbl(tblks)) > (cDbl(ablks)) then%>
								<td  class="big align="center"><%=round(diffblks,1)%></td>
						<% elseif (cDbl(tblks)) = (cDbl(ablks)) then %>
								<td  class="big align="center">E</td>
						<% end if %>
						
						<% if tasts < aasts then %>
								<td  class="big align="center">+<%=round(diffasts,1)%></td>
						<% elseif tasts > aasts then %>
								<td  class="big align="center"><%=round(diffasts,1)%></td>
						<% elseif tasts = aasts then %>
								<td  class="big align="center">E</td>
						<% end if %>
						
						<% if trebs < arebs then %>
								<td  class="big align="center">+<%=round(diffrebs,1)%></td>
						<% elseif trebs >arebs then %>
								<td  class="big align="center"><%=round(diffrebs,1)%></td>
						<% elseif trebs = arebs then %>
								<td  class="big align="center">E</td>
						<% end if %>
						
						<% if (cDbl(tpts)) < (cDbl(apts)) then %>
								<td class="big  align="center">+<%=round(diffpts,1)%></td>
						<% elseif (cDbl(tpts)) > (cDbl(apts)) then %>
								<td class="big  align="center"><%=round(diffpts,1)%></td>
						<% elseif (cDbl(tpts)) = (cDbl(apts))then %>
								<td  class="big align="center">E</td>
						<% end if %>

						<% if tstls < astls then %>
								<td  class="big align="center">+<%=round(diffstls,1)%></td>
						<% elseif tstls > astls then %>
								<td  class="big align="center"><%=round(diffstls,1)%></td>
						<% elseif tstls = astls then %>
								<td  class="big align="center">E</td>
						<% end if %>

						<% if tthrees < athrees then %>
								<td  class="big align="center">+<%=round(diffthrees,1)%></td>
						<% elseif tthrees > athrees then %>
								<td  class="big align="center"><%=round(diffthrees,1)%></td>
						<% elseif tthrees = athrees then %>
								<td class="big  align="center">E</td>
						<% end if %>
						
						<% if ttos < atos then %>
								<td class="big  align="center"><%=round(difftos,1)%></td>
						<% elseif ttos > atos then%>
								<td  class="big align="center"><%=round(difftos,1)%></td>
						<% elseif ttos = atos then%>		
								<td  class="big align="center">E</td>
						<% end if %>
						</tr>
						<tr>
						<td colspan="7" style="background-color:#FFEB3B;" align="center"><span style="font-weight:bold;color:black;" class="text-uppercase big"><i class="fa fa-exclamation-triangle" aria-hidden="true"></i>&nbsp;Net result of this Trade is <span> 
						<% if tbarps < abarps then %>
								<span class="badgeUp" data-toggle="tooltip" title="Barps Diff">+<%=round(diffBarps,1)%></span> Barps!
						<% elseif tbarps > abarps then  %>
								<span class="badgeDown" data-toggle="tooltip" title="Barps Diff"><%=round(diffBarps,1)%></span> Barps!
						<% elseif tbarps = abarps then %>
								<span class="badgeEven" data-toggle="tooltip" title="Barps Diff">E</span></span> Barps!
						<% end if %>
						</td>
						</tr>
						<tr style="text-align: center;">
							<td class="big" colspan="7">
							<small><strong><black>remain open for:</black></strong></small>
							<select name="expireTime" class="select">
							<option value="1">1 Day</option>
							<option value="2" selected="selected">2 Days</option>
							<option value="3">3 Days</option>
							<option value="4">4 Days</option>
							<option value="5">5 Days</option>
							<option value="6">6 Days</option>
							<option value="7">7 Days</option>
							</select>
							</td>						
						</tr>
						<tr>
									<td colspan="8"><textarea name="txtNotes" class="form-control" rows="3" placeholder="Comments" id="txtNotes"></textarea>
						</tr>
						<tr>
							<td colspan="8">
							<button type="submit" value="Submit Trade Offer" name="Action" class="btn btn-default-green  btn-block btn-md"><span class="glyphicon glyphicon-save"></span>&nbsp;Submit Trade Offer</button>
							<button type="submit" value="Forecast" name="Action" class="btn btn-default-blue btn-block btn-md"><i class="fa fa-bar-chart" aria-hidden="true"></i>&nbsp;Forecast Trade</button>
							<button type="submit" value="Counter" name="Action" class="btn btn-default btn-block btn-md"><i class="fas fa-edit"></i>&nbsp;Revise Trade</button>
							<button type="submit" value="Delete Invalid Trade" name="Action" class="btn btn-default-red btn-block btn-md"><i class="fas fa-trash-alt"></i>&nbsp;Delete Trade </button></button>
							</td>						
						</tr>
					</table>
				</div>
			</div>
	</div>	
</div>
</form>
<%
 	objRS.MoveNext
	
	diffBarps         = 0
	diffblks          = 0
	diffasts          = 0
	diffrebs          = 0       
	diffpts           = 0
	diffstls          = 0
	diffstls          = 0
	difftos           = 0
	diffthrees        = 0
	atotRegGameCnt    = 0
	ttotRegGameCnt    = 0
	atotPOGameCnt     = 0
	ttotPOGameCnt     = 0
	analysisRegGameCnt= 0
	analysisPOGameCnt = 0
	t2pid             = 0
	a2pid             = 0
	t3pid             = 0
	a3pid             = 0 
	
	t1first  = objRS.Fields("t1first").Value
	t1last   = objRS.Fields("t1last").Value
	t1pid    = objRS.Fields("t1PID").Value

	if objRS.Fields("t2PID").Value > 0 then
		t2first= objRS.Fields("t2first").Value
		t2last = objRS.Fields("t2last").Value
		t2pid  = objRS.Fields("t2PID").Value 
	end if

	if objRS.Fields("t3PID").Value > 0 then
		t3first= objRS.Fields("t3first").Value
		t3last = objRS.Fields("t3last").Value
		t3pid  = objRS.Fields("t3PID").Value 
	end if

	a1first  = objRS.Fields("a1first").Value
	a1last   = objRS.Fields("a1last").Value
	a1pid    = objRS.Fields("a1PID").Value

	if objRS.Fields("a2PID").Value > 0 then
		a2first= objRS.Fields("a2first").Value
		a2last = objRS.Fields("a2last").Value
		a2pid  = objRS.Fields("a2PID").Value
	end if

	if objRS.Fields("a3PID").Value > 0 then
		a3first= objRS.Fields("a3first").Value
		a3last = objRS.Fields("a3last").Value
		a3pid  = objRS.Fields("a3PID").Value
	end if

	Wend
%>
<%
end if
if sAction = "Submit Trade Offer" and errorcode = "Invalid Offer" then %>
<!--#include virtual="Common/headerMain.inc"-->
<form action="pendingAnalyzedTrades.asp" method="POST" name="frminvalidtrade" language="JavaScript">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="cmbTeam" value="<%= sTeam %>" />
<input type="hidden" name="var_tradepartneremail" value="<%=objRSTeam2.Fields("HomeEmail").Value%>" />
<input type="hidden" name="var_owneremail" value="<%=objRSTeam1.Fields("HomeEmail").Value%>" />
<input type="hidden" name="var_teamname" value="<%=objRSTeam1.Fields("TeamName").Value%>" />
<input type="hidden" name="var_playercnt" value="<%=objRSTeam1.Fields("ActivePlayerCnt").Value%>" />
<input type="hidden" name="var_tradepartnername" value="<%=objRSTeam2.Fields("teamname").Value%>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
				<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> See Below for Possible Errors
			</div>
			<ol>
					<li>The following trade offer can' be processed due to an invalid browser state.</li>
					<li>Exact trade is currently pending</li>
					<li>Players in this trade on not on rosters</li>
			</ol>
		</div>
	</div>
</div>
 <div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
			<A HREF="" onClick="history.back();return false;">
				<button type="submit" value="Cancel" formnovalidate  name="Cancel" class="btn btn-block btn-default btn-md"><i class="fa fa-hand-o-left" aria-hidden="true"></i>&nbsp;Cancel</button>
			</A> 
		</div>
	</div>
</div>
<%
End if
if sAction = "Forecast" or sAction = "Reforecast" then 
   GetAnyParameter "var_tradeid", stradeid
   tradeid        = Request.Form("var_tradeid")
	 'Response.Write " trade id value  = "&tradeid&" <br> "
	 
	 if sAction = "Forecast" then 
		wNeutralVal = false
	 else
		wNeutralVal = true
	 end if	

	 
	 objrstrade.Open "SELECT t.* FROM tblTradeAnalysis t WHERE t.tradeid = " & tradeid, objConn,3,3,1'						
   tplayer1       = objrstrade.Fields("TradedPlayerID").Value
   tplayer2       = objrstrade.Fields("TradedPlayerID2").Value
   tplayer3       = objrstrade.Fields("TradedPlayerID3").Value
   aplayer1       = objrstrade.Fields("AcquiredPlayerID").Value
   aplayer2       = objrstrade.Fields("AcquiredPlayerID2").Value
   aplayer3       = objrstrade.Fields("AcquiredPlayerID3").Value 
   TradeOwnerId   = objrstrade.Fields("ToOID").Value
   objrstrade.Close
   
   '###################################################################################################################
   ' trade_analysis.inc was written for TradeOffers coming to you.  Because of this we need to flip the variables for 
   ' trades that we are making so the display will be right. 
   '###################################################################################################################
   if tplayer1 <> 0 then FuncCall = Get_One_Name(tplayer1, saPlayer1Name, x, saRegGameCnt1, saPOGameCnt1) else saPlayer1Name = "" end if
   if tplayer2 <> 0 then FuncCall = Get_One_Name(tplayer2, saPlayer2Name, x, saRegGameCnt2, saPOGameCnt2) else saPlayer2Name = "" end if
   if tplayer3 <> 0 then FuncCall = Get_One_Name(tplayer3, saPlayer3Name, x, saRegGameCnt3, saPOGameCnt3) else saPlayer3Name = "" end if
   if aplayer1 <> 0 then FuncCall = Get_One_Name(aplayer1, stPlayer1Name, x, stRegGameCnt1, stPOGameCnt1) else stPlayer1Name = "" end if
   if aplayer2 <> 0 then FuncCall = Get_One_Name(aplayer2, stPlayer2Name, x, stRegGameCnt2, stPOGameCnt2) else stPlayer2Name = "" end if
   if aplayer3 <> 0 then FuncCall = Get_One_Name(aplayer3, stPlayer3Name, x, stRegGameCnt3, stPOGameCnt3) else stPlayer3Name = "" end if
   
   wPositive = 0
   wNegative = 0
   wEven = 0
   
	 if wNeutralVal = false then 
		 objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'NEUTRAL_VALUE' ",objConn
		 wNeutralVal = objRSWork.Fields("param_amount").Value
		 objRSWork.Close
		 'Response.Write " Hitting the (IF) initial Neutral Value from the Table = "&wNeutralVal&" <br> "	
	 else
		 wNeutralVal = cint(Request.Form("reForcast"))
		 'Response.Write " Hitting the (ELSE) Neutral Value from the Screen = "&wNeutralVal&" <br> "
	 end if
   
   objRSAll.Open      	"SELECT * FROM tblGameDeadLines where gameday >= date() order by gameday", objConn					
   While Not objRSAll.EOF
		gameday  = objRSAll("gameday")	
	  wRetcd   = Forecast_Lineup(gameday,ownerid,0,CenName,CenPID,CenBarps,For1Name,For1PID,For1Barps,For2Name,For2Pid,For2Barps,Gua1Name,Gua1PID,Guard1Barps,Gua2Name,Gua2PID,Guard2Barps)
	  wRetcd   = Forecast_Lineup(gameday,ownerid,TradeOwnerId,OppCenName,OppCenPID,OppCenBarps,OppFor1Name,OppFor1PID,OppFor1Barps,OppFor2Name,OppFor2PID,OppFor2Barps,OppGua1Name,OppGua1PID,OppGuard1Barps,OppGua2Name,OppGua2PID,OppGuard2Barps)	
	  
	  myTotal  = cDbl(CenBarps)+cDbl(For1Barps)+cDbl(For2Barps)+cDbl(Guard1Barps)+cDbl(Guard2Barps)
	  OppTotal = cDbl(OppCenBarps)+cDbl(OppFor1Barps)+cDbl(OppFor2Barps)+cDbl(OppGuard1Barps)+cDbl(OppGuard2Barps)
	  
	  xdiff = MyTotal - OppTotal
	  if xdiff < 0 then
	     xdiff = xdiff * -1
	  end if
      	  
	  if xdiff <= wNeutralVal then
	     wEven = wEven + 1
	  elseif OppTotal > MyTotal then
	     wPositive = wPositive + 1
	  else
	     wNegative = wNegative + 1
	  end if
	  
	  'if OppTotal > MyTotal then
	  '   wPositive = wPositive + 1
	  'elseif MyTotal > OppTotal then
	  '   wNegative = wNegative + 1
	  'else
	  '   wEven = wEven + 1
	  'end if
	  
	  'Response.Write "gameday="&gameday&","&CenBarps&","&For1Barps&","&For2Barps&","&Guard1Barps&","&Guard2Barps&","
	  'Response.Write "myTotal="&myTotal&", OppTotal="&OppTotal&", wPositive= "&wPositive&", wNegative= "&wNegative&", wEven="&wEven&"<br>"
      objRSAll.MoveNext
   Wend
   objRSAll.Close	  

%>
<form action="pendingAnalyzedTrades.asp" name="frmtradeviolation" method="POST" onSubmit="return processReforecast(this)">
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("towner").Value %>" />
  <input type="hidden" name="var_teamname" value="<%=objRS.Fields("tteamname").Value %>" />
  <input type="hidden" name="var_myname" value="<%=objRS.Fields("ateamname").Value %>" />
  <input type="hidden" name="var_tradedeleted" value="<%= tradedeleted %>" />
	<input type="hidden" name="var_tradeid" value="<%=objRS.Fields("tradeid").Value %>" />

  <!--#include virtual="Common/headermain.inc"-->
  <!--#include virtual="Common/trade_analysis.inc"-->
	<%
	Set	objRSHurt       = Server.CreateObject("ADODB.RecordSet")					 
	objRSHurt.Open "SELECT * from tblplayers where IR = true and ownerID =  " & ownerid & " order by firstname",objConn,3,3,1	
	%>

 <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">
	<%if objRSHurt.RecordCount > 0 then %>
	<table class="table table-custom-black table-bordered table-condensed">
		<tr style="background-color:#9a1400;vertical-align:middle">
		<td>
			<table class="table table-striped table-responsive table-custom table-condensed">
			<td style="text-align:left;">	<i class="fas fa-briefcase-medical red"></i>&nbsp;Your Injured Players are Omitted from Forecast. Visit the Player Profile Page by Clicking the link to set the Indicator to Off.</td>
			<%
			While Not objRSHurt.EOF
			%>
				<tr>
						<td><a class="blue" href="playerprofile.asp?pid=<%=objRSHurt.Fields("PID").Value %>"><%=objRSHurt.Fields("firstName").Value%>&nbsp;<%=objRSHurt.Fields("LastName").Value%></a></td>
				</tr>
			<%
			objRSHurt.MoveNext
			Wend
			%>
			</table>
		</td>
		</tr>
	</table>
	<br>
		 </div>
		</div>
  </div>
	<%end if%>
	 <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">
	<%

objRSAll.Open      	"SELECT * FROM tblGameDeadLines where gameday >= date() order by gameday", objConn					

While Not objRSAll.EOF

	gameday       = objRSAll("gameday")
	CenName       = "<redText>No Center</redText>" 
	For1Name      = "<redText>No Forward</redText>"
	For2Name      = "<redText>No Forward</redText>"
	Gua1Name    	= "<redText>No Guard</redText>"
	Gua2Name    	= "<redText>No Guard</redText>"
	OppCenName    = "<redText>No Center</redText>"
	OppFor1Name   = "<redText>No Forward</redText>"
	OppFor2Name   = "<redText>No Forward</redText>"
	OppGua1Name 	= "<redText>No Guard</redText>"
	OppGua2Name 	= "<redText>No Guard</redText>"		

	'Response.Write "tplayer1="&tplayer1&", tplayer2="&aplayer2&", tplayer3="&tplayer3&", aplayer1="&aplayer1&", aplayer2="&aplayer2&", aplayer3="&aplayer3&"  <br>" 
	
	'Call the first time with 0 in the TraderId position
	'Call the second time with the ownerid of the Person proposing the trade.  Need a variable for 2nd call.  Hardcoded now.
	
	
	wRetcd = Forecast_Lineup(gameday,ownerid,0,CenName,CenPID,CenBarps,For1Name,For1PID,For1Barps,For2Name,For2Pid,For2Barps,Gua1Name,Gua1PID,Guard1Barps,Gua2Name,Gua2PID,Guard2Barps)
	wRetcd = Forecast_Lineup(gameday,ownerid,TradeOwnerId,OppCenName,OppCenPID,OppCenBarps,OppFor1Name,OppFor1PID,OppFor1Barps,OppFor2Name,OppFor2PID,OppFor2Barps,OppGua1Name,OppGua1PID,OppGuard1Barps,OppGua2Name,OppGua2PID,OppGuard2Barps)	
	
	if IsNull(CenBarps)       then CenBarps       = 0 end if
	if IsNull(For1Barps)      then For1Barps      = 0 end if
	if IsNull(For2Barps)      then For2Barps      = 0 end if
	if IsNull(Guard1Barps)    then Guard1Barps    = 0 end if
	if IsNull(Guard2Barps)    then Guard2Barps    = 0 end if
	if IsNull(OppCenBarps)    then OppCenBarps    = 0 end if
	if IsNull(OppFor1Barps)   then OppFor1Barps   = 0 end if
	if IsNull(OppFor2Barps)   then OppFor2Barps   = 0 end if
	if IsNull(OppGuard1Barps) then OppGuard1Barps = 0 end if
	if IsNull(OppGuard2Barps) then OppGuard2Barps = 0 end if

	myTotal  = cDbl(CenBarps)+cDbl(For1Barps)+cDbl(For2Barps)+cDbl(Guard1Barps)+cDbl(Guard2Barps)
	OppTotal = cDbl(OppCenBarps)+cDbl(OppFor1Barps)+cDbl(OppFor2Barps)+cDbl(OppGuard1Barps)+cDbl(OppGuard2Barps)
	
	xdiff = MyTotal - OppTotal
	if xdiff < 0 then
	     xdiff = xdiff * -1
	end if
	
	'#########################################################
	'Get the Opponent Name
	'You have to use 3,3,1 if you want to get the record count.
	'If recordcount is not = 1 then set Opponent = "TBD"
	'#########################################################
    objRSOpponentName.Open 	"SELECT * FROM qryAllGames where (AwayTeamInd = " & ownerid & " OR HomeTeamInd  = " & ownerid & ") AND  gameday = #" & gameday & "#", objConn,3,3,1
	if objRSOpponentName.RecordCount = 1 then
		if objRSOpponentName("HomeTeamInd") = ownerid then
		   OpponentName = objRSOpponentName("AwayTeamShort")
		else
		   OpponentName = objRSOpponentName("AwayTeamShort")
		end if	
	else
		OpponentName = "TBD"
	end if					
	
	objRSOpponentName.Close	

%>


			<!-- Trade Forecaster Logic-->
					<table class="table table-custom-black table-bordered table-condensed table-responsive">
						<tr> 
							<th class="text-center"><%=(FormatDateTime(objRSAll("gameday"),1))%></th>
						</tr>
					</table>
				<table class="table box table-custom-black table-bordered table-condensed">
					<tr>
						<th class="big" width="30%">Before</th>
						<th class="big" style="text-align:center;width:15%;">Avg</th>
						<th class="text-center big" style="vertical-align:middle;width:10%;">POS</th>		
						<th class="big" style="text-align:center;width:15%;">Avg</th>
						<th class="big" width="30%">After</th>
					</tr>
					<tr bgcolor="#FFFFFF">
					<% if (CenPID = tplayer1 or CenPID = tplayer2 or CenPID = tplayer3) and CenPID <> 0  then %>
						<td class="big" style="text-align: left"><blueText><%=CenName%></blueText></td>
					<% else %>
						<td class="big" style="text-align: left"><%=CenName%></td>
					<% end if %>		
						<td class="big " style="text-align: center"><%=round(CenBarps,2)%></td>
						<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">CEN</th>	
						<td class="big " style="text-align: center"><%=round(OppCenBarps,2)%></td>					
					<% if (OppCenPID = aplayer1 or OppCenPID = aplayer2 or OppCenPID = aplayer3) and OppCenPID <> 0  then %>
						<td class="big" style="text-align: left"><blueText><%=OppCenName%></blueText></td>
					<% else %>
						<td class="big" style="text-align: left"><%=OppCenName%></td>
					<% end if %>
					</tr>  

					<tr bgcolor="#FFFFFF">

					<% if (For1PID = tplayer1 or For1PID = tplayer2 or For1PID = tplayer3) and For1PID <> 0 then %>
						<td class="big" style="text-align: left"><blueText><%=For1Name%></blueText></td>
					<% else %>
						<td class="big" style="text-align: left"><%=For1Name%></td>
					<% end if %>	
					<td class="big "style="text-align: center"><%=round(For1Barps,2)%></td>		
						<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">FOR</th>						
					<td class="big " style="text-align: center"><%=round(OppFor1Barps,2)%></td>
					<% if (OppFor1PID = aplayer1 or OppFor1PID = aplayer2 or OppFor1PID = aplayer3) and OppFor1PID <> 0 then %>
							<td class="big" style="text-align: left"><blueText><%=OppFor1Name%></blueText></td>
					<% else %>
							<td class="big" style="text-align: left"><%=OppFor1Name%></td>
					<% end if %>					

					</tr>  
					
					<tr bgcolor="#FFFFFF">
					<% if (For2PID = tplayer1 or For2PID = tplayer2 or For2PID = tplayer3) and For2PID <> 0   then %>
							<td class="big" style="text-align: left"><blueText><%=For2Name%></blueText></td>
					<% else %>
							<td class="big" style="text-align: left"><%=For2Name%></td>
					<% end if %>					
					<td class="big " width="10%" style="text-align: center"><%=round(For2Barps,2)%></td>
					<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">FOR</th>	
					<td class="big " width="10%" style="text-align: center"><%=round(OppFor2Barps,2)%></td>
					<% if (OppFor2PID = aplayer1 or OppFor2PID = aplayer2 or OppFor2PID = aplayer3) and OppFor2PID <> 0   then %>
							<td class="big" style="text-align: left"><blueText><%=OppFor2Name%></blueText></td>
					<% else %>
							<td class="big" style="text-align: left"><%=OppFor2Name%></td>
					<% end if %>					
					</tr>  
					
					<tr bgcolor="#FFFFFF">
					<% if (Gua1PID = tplayer1 or Gua1PID = tplayer2 or Gua1PID = tplayer3) and Gua1PID <> 0  then %>
							<td class="big" style="text-align: left"><blueText><%=Gua1Name%></blueText></td>
					<% else %>
							<td class="big" style="text-align: left"><%=Gua1Name%></td>
					<% end if %>					
					<td class="big " width="10%" style="text-align: center"><%=round(Guard1Barps,2)%></td>
					<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">GUA</th>	
					<td class="big " width="10%" style="text-align: center"><%=round(OppGuard1Barps,2)%></td>
					<% if (OppGua1PID = aplayer1 or OppGua1PID = aplayer2 or OppGua1PID = aplayer3) and OppGua1PID <> 0  then %>
							<td class="big" style="text-align: left"><blueText><%=OppGua1Name%></blueText></td>
					<% else %>
							<td class="big" style="text-align: left"><%=OppGua1Name%></td>
					<% end if %>
					</tr>  
					
					<tr bgcolor="#FFFFFF">
					<% if (Gua2PID = tplayer1 or Gua2PID = tplayer2 or Gua2PID = tplayer3) and Gua2PID <> 0 then %>
						<td class="big" style="text-align: left"><blueText><%=Gua2Name%></blueText></td>
					<% else %>
						<td class="big" style="text-align: left"><%=Gua2Name%></td>
					<% end if %>					
					<td class="big " width="10%" style="text-align: center"><%=round(Guard2Barps,2)%></td>
					<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">GUA</th>	
					<td class="big " width="10%" style="text-align: center"><%=round(OppGuard2Barps,2)%></td>
					<% if (OppGua2PID = aplayer1 or OppGua2PID = aplayer2 or OppGua2PID = aplayer3) and OppGua2PID <> 0 then %>
						<td class="big" style="text-align: left"><blueText><%=OppGua2Name%></blueText></td>
					<% else %>
						<td class="big" style="text-align: left"><%=OppGua2Name%></td>
					<% end if %>
					</tr>  
					<tr style="background-color:white;text-align:center;vertical-align:middle;font-weight: bold;">
						<td class="big text-right">Totals&nbsp;<span style="text-align: right;"><i class="fal fa-arrow-to-right red"></i></span></td>
						<td class="big"><%= round(myTotal,2)%></td>
						<td class="big">
							<%if xdiff <= wNeutralVal then%>
							<evenIcon><i class="fa fa-balance-scale" aria-hidden="true"></i></evenIcon>				
							<%elseif OppTotal > myTotal then%>
							<greenIcon><i class="fa fa-thumbs-up" aria-hidden="true"></i></greenIcon>
							<%else%>	
							<redIcon><i class="fa fa-thumbs-down" aria-hidden="true"></i></redIcon>				
							<%end if %>
						</td>
						<td class="big"><%= round(OppTotal,2)%></td>
						<td class="big text-left"><span style="text-align: left;"><i class="fal fa-arrow-to-left red"></i></span>&nbsp;Totals</td>
					</tr>			
			</table>
			 <br>				
				<%
					objRSAll.MoveNext
					Wend
					objRSAll.Close	
									  			
				%>					

      </div>
			<!-- End Trade Forecaster Logic-->
    </div>
 </div>

 <div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
			<A HREF="" onClick="history.back();return false;">
				<button type="submit" value="Cancel" formnovalidate  name="Cancel" class="btn btn-default  btn-md"><i class="fa fa-hand-o-left" aria-hidden="true"></i>&nbsp;Back</button>
			</A> 
		</div>
	</div>
</div>
<br>
</form>
<% end if%>
</body>
<%
  objConn.Close
  Set objConn = Nothing
%>
</html>
