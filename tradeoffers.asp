<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<script language="JavaScript" type="text/javascript">
</script>
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn,objRS,ownerid,sAction,tradeid,errorcode,errormessage,tradername,myname,errormessage1,tradedeleted,tradecountered
	Dim tplayer1,tplayer2,tplayer3,aplayer1,aplayer2,aplayer3,tplayercnt,aplayercnt,objrsateam,objrstteam
	Dim objRSteam, objrsNames, objrstrade, w_email,cmbteam

	'ENAIL VARIABLES
	Dim email_to, email_subject, host, username, password, reply_to, port, from_address
	Dim first_name, last_name, home_address, email_from, telephone, comments, error_message
	Dim ObjSendMail, email_message, objEmail

	dim abbg, aapg, arpg, aspg, atpg, a3pg, abpg
	dim tbbg, tapg, trpg, tspg, ttpg, t3pg, tbpg
	
	GetAnyParameter "Action", sAction
	GetAnyParameter "var_tradeid", stradeid

	Set objConn            = Server.CreateObject("ADODB.Connection")
	Set objRS              = Server.CreateObject("ADODB.RecordSet")
	Set objRStteam         = Server.CreateObject("ADODB.RecordSet")
	Set objRSateam         = Server.CreateObject("ADODB.RecordSet")
	Set objRSteam          = Server.CreateObject("ADODB.RecordSet")
	Set objrsNames         = Server.CreateObject("ADODB.RecordSet")
	Set objrstrade         = Server.CreateObject("ADODB.RecordSet")
	Set objRSNext5         = Server.CreateObject("ADODB.RecordSet")
	Set objRSPO            = Server.CreateObject("ADODB.RecordSet")
	Set objRSlineups       = Server.CreateObject("ADODB.RecordSet")
	Set objRSWork          = Server.CreateObject("ADODB.RecordSet")
	
	'START FORECASTER LOGIC
	Dim OpponentName 
	Dim objRSHome, objRSAway, objRSAll,objRSPlayers, objRSOpponentName
	
	Set objRSHome        = Server.CreateObject("ADODB.RecordSet")
	Set objRSAway        = Server.CreateObject("ADODB.RecordSet")
	Set objRSAll         = Server.CreateObject("ADODB.RecordSet")
	Set objRSPlayers     = Server.CreateObject("ADODB.RecordSet")
	Set objRSOpponentName= Server.CreateObject("ADODB.RecordSet")
	Set objRSStaggerClose= Server.CreateObject("ADODB.RecordSet") 
	Set objEmail         = Server.CreateObject("ADODB.RecordSet") 

	'END FORECASTER LOGIC
 	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
									"Data Source=lineupstest.mdb;" & _
									"Persist Security Info=False"

	%>
	<!--#include virtual="Common/session.inc"-->
	<%
	tradeid = stradeid
	myname = Request.Form("var_myname")
	'Response.Write "Trade ID: " & tradeid & "<br> "


	objrstrade.Open "SELECT t.*, t2.TeamName " & _
	                "FROM tblpendingtrades t, tblowners t2 " & _
	                "WHERE t.tradeid = " & tradeid & "  "  & _
									"and  t.Fromoid = t2.ownerid ", objConn,3,3,1

	w_trade_count = objrstrade.Recordcount
	'Response.Write "Trade Count: " & w_trade_count& "<br> "
	'Response.Write "SQL: " & strSQL & "<br> "


	tplayer1  = objrstrade.Fields("TradedPlayerID").Value
	tplayer2  = objrstrade.Fields("TradedPlayerID2").Value
	tplayer3  = objrstrade.Fields("TradedPlayerID3").Value
	aplayer1  = objrstrade.Fields("AcquiredPlayerID").Value
	aplayer2  = objrstrade.Fields("AcquiredPlayerID2").Value
	aplayer3  = objrstrade.Fields("AcquiredPlayerID3").Value
	towner    = objrstrade.Fields("FromOid").Value
	tradername= objrstrade.Fields("TeamName").Value
	myownerID =  objrstrade.Fields("ToOid").Value

	tplayercnt  = 0
	aplayercnt  = 0

	if tplayer1 <> 0 then
		tplayercnt= tplayercnt + 1
	end if

	if tplayer2 <> 0 then
		tplayercnt= tplayercnt + 1
	end if

	if tplayer3 <> 0 then
		tplayercnt= tplayercnt + 1
	end if

	if aplayer1 <> 0 then
		aplayercnt= aplayercnt + 1
	end if

	if aplayer2 <> 0 then
		aplayercnt= aplayercnt + 1
	end if

	if aplayer3 <> 0 then
		aplayercnt= aplayercnt + 1
	end if


	objrstrade.close

select case sAction
	
	case "Forecast"
	conAction = "Continue"
	
	var_tradepartner = Request.Form("var_tradepartner")
	var_ownerid = Request.Form("var_ownerid")
	var_tradeid = Request.Form("var_tradeid")
	
	case "Counter"
	sURL = "tradeanalyzer.asp"
	conAction = "Continue"
	
	var_tradepartner = Request.Form("var_tradepartner")
	var_ownerid = Request.Form("var_ownerid")
	var_tradeid = Request.Form("var_tradeid")
	
	AddLinkParameter "var_ownerid", var_ownerid, sURL	
	AddLinkParameter "Action", conAction, sURL
  AddLinkParameter "cmbTeam", var_tradepartner, sURL
	AddLinkParameter "var_tradeid", var_tradeid, sURL
  	
	Response.Redirect sURL

	case "Decline"
	dim strMsgInfo

	tradeNotes = Request.Form("txtNotes")
	FuncCall = Get_Player_Names(tradeid, msgtradeplayers, msgacquireplayers)

	strSQL = "DELETE FROM tblpendingtrades WHERE TradeID =  "& tradeid & ";"
	objConn.Execute strSQL
	tradedeleted = "yes"

	FuncCall = Reset_PendTrade_Flags

	objRSWork.Open "SELECT shortName FROM tblOwners where ownerid = "&myownerID, objConn
	myShort = objRSWork.Fields("shortName").Value
	objRSWork.Close
	
	wEmailOwnerID  = towner
	wAlert         = "receiveTradeAlerts"
	email_subject  = "Trade Declined by " & myShort
	email_message = msgtradeplayers & "<br>"
	email_message = email_message & "for <br><br>"
	email_message = email_message & msgacquireplayers
	if len(tradeNotes) > 0 then
	   email_message = email_message&"<br>"&tradeNotes
	end if
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%			

		dim TradeCnt,objRSTrades
		
		Set objRSTrades  = Server.CreateObject("ADODB.RecordSet")	
		objRSTrades.Open "SELECT * FROM qryPendingTrade WHERE (((qryPendingTrade.aowner)=" & ownerid & ")) order by DecisionDate", objConn,3,3,1
		TradeCnt   = objRSTrades.RecordCount
	  'Response.Write "Record Count = : " & analysisCnt  & "  <br>"
		
		if TradeCnt > 0 then
			sURL = "tradeoffers.asp"
		else
			sURL = "dashboard.asp"
		end if	
		
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL
	    
	case "Accept"
	
	   objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE_DEADLINE' ",objConn
	   wTradeDeadLine = objRSWork.Fields("param_date").Value
	   objRSWork.Close
	
  	   if (wTradeDeadLine + 1 + 1/24) < now() then
		  errorcode   = "Trade Deadline Passed"
	   else
	      errorcode = "Confirmation"		
	   end if
%>				
<!--#include virtual="Common/setwaiversall.inc"-->
<%		
	case "Trade Confirmation"		
		errorcode = "False"
		if w_trade_count <= 0 then
			'response.write " Trade Deleted <br>"
			errorcode = "Trade Deleted"
		else				    

			objrsateam.Open	"SELECT * FROM tblplayers WHERE tblplayers.ownerid =" & ownerid & " and rentalPlayer = 0", objConn,3,3,1
			objrstteam.Open	"SELECT * FROM tblplayers WHERE tblplayers.ownerid =" & towner & " and rentalPlayer = 0", objConn,3,3,1
			
			w_new_accept_ct = (objrsateam.Recordcount - aplayercnt) + tplayercnt
			w_new_trader_ct = (objrstteam.Recordcount - tplayercnt) + aplayercnt

			'Response.Write "w_new_accept_ct = "&w_new_accept_ct&", w_new_trader_ct = "&w_new_trader_ct&"<br>"
			'Response.Write "aplayer1 = "&aplayer1&", aplayer2 = "&aplayer2&", aplayer3 = "&aplayer3&"<br>"
			'Response.Write "tplayer1 = "&tplayer1&", tplayer2 = "&tplayer2&", tplayer3 = "&tplayer3&"<br>"
			
			objrsateam.Close
			objrstteam.Close
       
			if w_new_accept_ct > 14 then
				errorcode   = "Trade Violation"
				errormessage= "Processing this trade violates the Roster Limit of 14!"
			elseif w_new_trader_ct > 14 then
				errorcode   = "Trade Violation"
				errormessage= "Processing this trade violates the Roster Limit of 14!"
			end if
			
			'Response.Write "errorCode = "&errorCode&"<br>"
			
			if errorCode = "False" then			    
				objRSlineups.Open "SELECT * FROM tbl_lineups where gameday = date() AND ownerID in ("&ownerid&","&towner&") " & _
                                  "AND final_lineup = FALSE AND (" & _
					"(sCenter in ("&aplayer1&","&aplayer2&","&aplayer3&","&tplayer1&","&tplayer2&","&tplayer3&") AND sCenterTip < time() - 1/24) OR " & _
					"(sForward in ("&aplayer1&","&aplayer2&","&aplayer3&","&tplayer1&","&tplayer2&","&tplayer3&") AND sForwardTip < time() - 1/24) OR " & _
					"(sForward2 in ("&aplayer1&","&aplayer2&","&aplayer3&","&tplayer1&","&tplayer2&","&tplayer3&") AND sForwardTip2 < time() - 1/24) OR " & _
					"(sGuard in ("&aplayer1&","&aplayer2&","&aplayer3&","&tplayer1&","&tplayer2&","&tplayer3&") AND sGuardTip < time() - 1/24) OR " & _
					"(sGuard2 in ("&aplayer1&","&aplayer2&","&aplayer3&","&tplayer1&","&tplayer2&","&tplayer3&") AND sGuardTip2 < time() - 1/24) " & _
								       ") ",objConn,3,3,1
			
				if objRSlineups.RecordCount > 0 then
					errorcode = "In Play Violation"
					errormessage = "<br>"
					wCount = 0
					'Response.Write "Center = "&objRSlineups.Fields("sCenter").Value&"<br>"
					'Response.Write "Tip  = "&objRSlineups.Fields("sCenterTip").Value&"<br>"
					While Not objRSlineups.EOF
						if (aplayer1 = objRSlineups.Fields("sCenter").Value and objRSlineups.Fields("sCenterTip").Value < time() - 1/24)    OR _
							 (aplayer1 = objRSlineups.Fields("sForward").Value and objRSlineups.Fields("sForwardTip").Value < time() - 1/24)   OR _  
						   (aplayer1 = objRSlineups.Fields("sForward2").Value and objRSlineups.Fields("sForwardTip2").Value < time() - 1/24) OR _ 
						   (aplayer1 = objRSlineups.Fields("sGuard").Value and objRSlineups.Fields("sGuardTip").Value < time() - 1/24)       OR _ 
						   (aplayer1 = objRSlineups.Fields("sGuard2").Value and objRSlineups.Fields("sGuardTip2").Value < time() - 1/24)  then						   
							FuncCall = Get_One_Name(aplayer1, playerName, x, x, x)						
							errormessage = errormessage & playerName & "<br>"
						end if
						
						if (aplayer2 = objRSlineups.Fields("sCenter").Value and objRSlineups.Fields("sCenterTip").Value < time() - 1/24)    OR _
							 (aplayer2 = objRSlineups.Fields("sForward").Value and objRSlineups.Fields("sForwardTip").Value < time() - 1/24)   OR _  
						   (aplayer2 = objRSlineups.Fields("sForward2").Value and objRSlineups.Fields("sForwardTip2").Value < time() - 1/24) OR _ 
						   (aplayer2 = objRSlineups.Fields("sGuard").Value and objRSlineups.Fields("sGuardTip").Value < time() - 1/24)       OR _ 
						   (aplayer2 = objRSlineups.Fields("sGuard2").Value and objRSlineups.Fields("sGuardTip2").Value < time() - 1/24)  then						   
							FuncCall = Get_One_Name(aplayer2, playerName, x, x, x)						
							errormessage = errormessage & playerName & "<br>"
						end if						
						
						if (aplayer3 = objRSlineups.Fields("sCenter").Value and objRSlineups.Fields("sCenterTip").Value < time() - 1/24)    OR _
							 (aplayer3 = objRSlineups.Fields("sForward").Value and objRSlineups.Fields("sForwardTip").Value < time() - 1/24)   OR _  
						   (aplayer3 = objRSlineups.Fields("sForward2").Value and objRSlineups.Fields("sForwardTip2").Value < time() - 1/24) OR _ 
						   (aplayer3 = objRSlineups.Fields("sGuard").Value and objRSlineups.Fields("sGuardTip").Value < time() - 1/24)       OR _ 
						   (aplayer3 = objRSlineups.Fields("sGuard2").Value and objRSlineups.Fields("sGuardTip2").Value < time() - 1/24)  then						   
							FuncCall = Get_One_Name(aplayer3, playerName, x, x, x)						
							errormessage = errormessage & playerName & "<br>"
						end if

						if (tplayer1 = objRSlineups.Fields("sCenter").Value and objRSlineups.Fields("sCenterTip").Value < time() - 1/24)    OR _
							 (tplayer1 = objRSlineups.Fields("sForward").Value and objRSlineups.Fields("sForwardTip").Value < time() - 1/24)   OR _  
						   (tplayer1 = objRSlineups.Fields("sForward2").Value and objRSlineups.Fields("sForwardTip2").Value < time() - 1/24) OR _ 
						   (tplayer1 = objRSlineups.Fields("sGuard").Value and objRSlineups.Fields("sGuardTip").Value < time() - 1/24)       OR _ 
						   (tplayer1 = objRSlineups.Fields("sGuard2").Value and objRSlineups.Fields("sGuardTip2").Value < time() - 1/24)  then						   
							FuncCall = Get_One_Name(tplayer1, playerName, x, x, x)						
							errormessage = errormessage & playerName & "<br>"
						end if

						if (tplayer2 = objRSlineups.Fields("sCenter").Value and objRSlineups.Fields("sCenterTip").Value < time() - 1/24)    OR _
							 (tplayer2 = objRSlineups.Fields("sForward").Value and objRSlineups.Fields("sForwardTip").Value < time() - 1/24)   OR _  
						   (tplayer2 = objRSlineups.Fields("sForward2").Value and objRSlineups.Fields("sForwardTip2").Value < time() - 1/24) OR _ 
						   (tplayer2 = objRSlineups.Fields("sGuard").Value and objRSlineups.Fields("sGuardTip").Value < time() - 1/24)       OR _ 
						   (tplayer2 = objRSlineups.Fields("sGuard2").Value and objRSlineups.Fields("sGuardTip2").Value < time() - 1/24)  then						   
							FuncCall = Get_One_Name(tplayer2, playerName, x, x, x)						
							errormessage = errormessage & playerName & "<br>"
						end if
						
						if (tplayer3 = objRSlineups.Fields("sCenter").Value and objRSlineups.Fields("sCenterTip").Value < time() - 1/24)    OR _
							 (tplayer3 = objRSlineups.Fields("sForward").Value and objRSlineups.Fields("sForwardTip").Value < time() - 1/24)   OR _  
						   (tplayer3 = objRSlineups.Fields("sForward2").Value and objRSlineups.Fields("sForwardTip2").Value < time() - 1/24) OR _ 
						   (tplayer3 = objRSlineups.Fields("sGuard").Value and objRSlineups.Fields("sGuardTip").Value < time() - 1/24)       OR _ 
						   (tplayer3 = objRSlineups.Fields("sGuard2").Value and objRSlineups.Fields("sGuardTip2").Value < time() - 1/24)  then						   
							FuncCall = Get_One_Name(tplayer3, playerName, x, x, x)						
							errormessage = errormessage & playerName & "<br>"
						end if						
						
						objRSlineups.MoveNext						
					Wend	
					
				end if					    
				objRSlineups.close								
				
				objRSStaggerClose.Open "SELECT * FROM tbltimedEvents where event = 'setwaiversall' ",objConn
				wStaggerClose = FormatDateTime(objRSStaggerClose.Fields("nextrun_EST")-(1/24), 3)
				objRSStaggerClose.Close
								 					  
                errormessage = errormessage & "<br>Restructure this trade, or wait until the Stagger Window closes at <strong>"&wStaggerClose&"</strong> to accept!!!"					  
			end if
			
			
			if errorcode = "False" then
				'##################################################################################
				'UPDATE PLAYERS TABLE-SWITCH TRADE OWNERID (THE INITIATOR OF THE TRADE) TO NEW TEAM
				'##################################################################################
				strSQL = "UPDATE tblPlayers " &_
				         "SET ownerid = "&towner&", ontheblock = 0, ir = 0, PendingWaiver = 0 " &_
				         "WHERE PID in ("&aplayer1&", "&aplayer2&", "&aplayer3&") "
				objConn.Execute strSQL

				'#####################################################################################
				'UPDATE PLAYERS TABLE-SWITCH ACQUIRED OWNERID (THE RECEPIENT OF THE OFFER) TO NEW TEAM
				'######################################################################################
				strSQL = "UPDATE tblPlayers " &_
				         "SET ownerid = "&ownerid&", ontheblock = 0, ir = 0, PendingWaiver = 0 " &_
				         "WHERE PID in ("&tplayer1&", "&tplayer2&", "&tplayer3&") "
				objConn.Execute strSQL

				FuncCall = Get_Player_Names(tradeid, msgtradeplayers, msgacquireplayers)

				'DELETE ACCEPTED TRADE
				strSQL = "DELETE FROM tblpendingtrades WHERE TradeID = "&tradeid&" "
				objConn.Execute strSQL

				'Response.Write "Trade Processed, Check Roster for Verification"
                objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE' ",objConn
				TxnCost = objRSWork.Fields("param_amount").Value
				objRSWork.Close
				
				TransType = "Trade With"

				'*****************************************************
				'Process Transactions for Owner initiating the Trade 
				'*****************************************************
				nAddCt = 0
				nReleaseCt = 0
				if aplayer1 > 0 then
				    nAddCt = nAddCt + 1
					FuncCall = DelPendingTrades(aPlayer1)
				end if

				if aplayer2 > 0 then
					nAddCt = nAddCt + 1
					FuncCall = DelPendingTrades(aPlayer2)
				end if

				if aplayer3 > 0 then
					nAddCt = nAddCt + 1
					FuncCall = DelPendingTrades(aPlayer3)
				end if

				'*****************************************************
				'Process Transactions for Owner accepting the Trade
				'*****************************************************
         
				if tplayer1 > 0 then
					nReleaseCt = nReleaseCt + 1
					FuncCall = DelPendingTrades(tplayer1)
				end if

				if tplayer2 > 0 then
					nReleaseCt = nReleaseCt + 1
					FuncCall = DelPendingTrades(tplayer2)
				end if

				if tplayer3 > 0 then
					nReleaseCt = nReleaseCt + 1
					FuncCall = DelPendingTrades(tplayer3)
				end if
				
				TotCost = TxnCost * nAddCt
		        strSQL ="insert into tblTransactions " & _
				"(OwnerID,TransType,TransCost,TradedFrom,transAddPlayerCnt,transAddPlayer1,transAddPlayer2,transAddPlayer3," & _
				"transReleasePlayerCnt,transReleasePlayer1,transReleasePlayer2,transReleasePlayer3) " & _
		        "values ("&towner&",'"&TransType&"',"&TotCost&","&ownerid&","&nAddCt&","&aPlayer1&","&aPlayer2&","&aPlayer3&"," & _
				          nReleaseCt&","&tPlayer1&","&tPlayer2&","&tPlayer3&") "
		        objConn.Execute strSQL
		
				'Response.Write "Sql = " & strSQL  & "<br>" 
				
		        TotCost = TxnCost * nReleaseCt
				strSQL ="insert into tblTransactions " & _
				"(OwnerID,TransType,TransCost,TradedFrom,transAddPlayerCnt,transAddPlayer1,transAddPlayer2,transAddPlayer3," & _
				"transReleasePlayerCnt,transReleasePlayer1,transReleasePlayer2,transReleasePlayer3) " & _
		        "values ("&ownerid&",'"&TransType&"',"&TotCost&","&towner&","&nReleaseCt&","&tPlayer1&","&tPlayer2&","&tPlayer3&"," & _
				          nAddCt&","&aPlayer1&","&aPlayer2&","&aPlayer3&") "
		        objConn.Execute strSQL
				
				'Response.Write "Sql = " & strSQL  & "<br>" 
				
				FuncCall = Reset_PendTrade_Flags
				FuncCall = Remove_From_Lineups(aplayer1,aplayer2,aplayer3,tplayer1,tplayer2,tplayer3,ownerid,towner)

				'*****************************************************
				'Update Active Player Counts on the
				'*****************************************************
				'UPDATE Active Player count for accepting owner
				strSQL = "UPDATE tblOwners SET activePlayerCnt = "& w_new_accept_ct & "	WHERE ownerid =" & ownerid & ";"
				objConn.Execute strSQL

				'UPDATE Active Player count for initiating owner
				strSQL = "UPDATE tblOwners SET activePlayerCnt = "& w_new_trader_ct & "	WHERE ownerid =" & towner & ";"
				objConn.Execute strSQL
				
				objRSWork.Open "SELECT shortName FROM tblOwners where ownerid = "&towner, objConn
				traderShort = objRSWork.Fields("shortName").Value
				objRSWork.Close
				
				objRSWork.Open "SELECT shortName FROM tblOwners where ownerid = "&ownerid, objConn
				myShort = objRSWork.Fields("shortName").Value
				objRSWork.Close
				
				wEmailOwnerID  = null
				wAlert         = "receiveTradeAlerts"
				email_subject  = "Trade Accepted: "&traderShort&" & "&myShort
				email_message = msgtradeplayers & "<br>"
				email_message = email_message & "for <br><br>"
				email_message = email_message & msgacquireplayers
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%	
				
				sURL = "dashboard.asp"
				AddLinkParameter "var_ownerid", ownerid, sURL
				Response.Redirect sURL

			end if
		end if

	case ""
	 	'ownerid = Request.querystring("ownerid")
		ownerid = session("ownerid")	
	 	if ownerid = "" then
			GetAnyParameter "var_ownerid", ownerid
		end if

   case "My-IGBL"
		sURL = "dashboard.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
    Response.Redirect sURL

  case "Return"
    'ownerid = Request.querystring("ownerid")
		ownerid = session("ownerid")	
    case "Delete Invalid Trade"
		'Response.Write "Trade ID  = : " & stradeid  & "  <br>"
    strSQL = "DELETE FROM tblpendingtrades where tradeid =" & stradeid & " ;"
		objConn.Execute strSQL

		sURL = "dashboard.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL

	end select

	'***********************************************************
	' DelPendingTrades()
	' - Search the pending Trades table to remove
	' 	all trades involving the players that were traded here.  Those
	' 	other trades are no longer valid.  This function will also
	' 	set the Pending Trade Flag to False for all players whose
	' 	Pending Trade is being deleted.
	' - Also remove any pending rows from tblTradeAnalysis involving players that were traded.
	'***********************************************************
	Function DelPendingTrades(p_PlayerID)

		strSQL = "DELETE FROM tblpendingtrades " & _
           "WHERE  tradedplayerid      = "& p_PlayerID & "  "  & _
                  "or tradedplayerid2 = "& p_PlayerID & "   " & _
                  "or tradedplayerid3 = "& p_PlayerID & "   " & _
                  "or acquiredplayerid = "& p_PlayerID & "  " & _
                  "or acquiredplayerid2 = "& p_PlayerID & " " & _
                  "or acquiredplayerid3 = "& p_PlayerID & "  "
		objConn.Execute strSQL
		
		strSQL = "DELETE FROM tblTradeAnalysis " & _
           "WHERE  tradedplayerid      = "& p_PlayerID & "  "  & _
                  "or tradedplayerid2 = "& p_PlayerID & "   " & _
                  "or tradedplayerid3 = "& p_PlayerID & "   " & _
                  "or acquiredplayerid = "& p_PlayerID & "  " & _
                  "or acquiredplayerid2 = "& p_PlayerID & " " & _
                  "or acquiredplayerid3 = "& p_PlayerID & "  "
		objConn.Execute strSQL

		'*************************************************************************
		'Delete all entries from tblWaivers table where player_id = Player Claimed
		'*************************************************************************
   		
		strSQL = "DELETE from tblWaivers where PID_Waived = " & p_PlayerID & ";"
		objConn.Execute strSQL
	
	End Function
	
			

	'***********************************************************
	' Get_Player_Names()
	'***********************************************************
   	Function Get_Player_Names(p_tradeID, msgtradeplayers, msgacquireplayers)

		objrsNames.Open "SELECT * FROM qryPendingTrade Where TradeID = "&p_tradeID&" " , objConn

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
		playerName = left(objrsNames.Fields("firstName").Value,1) & ". " & objrsNames.Fields("lastName").Value
		lName = objrsNames.Fields("lastName").Value
        objrsNames.Close
		
		objRSWork.Open  "SELECT * FROM qryAllPlayerGameDays WHERE pid = " & playerID & " and gameday >= Date() " & _
						"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') order by gameday ", objConn,1,1					 

		playerGmCnt = objRSWork.RecordCount
		objRSWork.Close
		
		objRSWork.Open 		"select gameday from qryMissingLineup " & _
							"where pid = " & playerID & " " & _
							"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1  												
		
		playerPOGMCnt = objRSWork.RecordCount							
		objRSWork.Close
		
		'Response.Write ", playerName = "&playerName&"<br>"
	End Function	
	
	'***********************************************************
	' Forecast_Lineup()
	' The Values of p_TraderID tells me which type of Query I should perform
	'    - If value is 0 then my result set should reflect my current roster and return how my lineups would look before the trade
	'    - If the value is not zero then my result set should reflect how my lineups would look after the trade.
	'***********************************************************	
	Function Forecast_Lineup (p_gameday, p_OwnerID, p_TraderID, p_CenName, CEN, CEN_BARPS, p_For1Name, F1, F1_BARPS, p_For2Name, F2, F2_BARPS, p_Guard1Name, G1, G1_BARPS, p_Guard2Name, G2, G2_BARPS)

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
								"(ownerID = "& p_OwnerID & " AND PID not in (" & aplayer1 & "," & aplayer2 & "," & aplayer3 & ") AND IR = 0 and Injury = 0 ) " & _
								"OR " &_									 
								"(ownerID = "& p_TraderID & " AND PID in (" & tplayer1 & "," & tplayer2& "," & tplayer3 & ") ) " & _									 
                                "   ) " , objConn,3,3,1		 
		 end if
		 
			 While Not objRSPlayers.EOF
						   
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

	t2pid = 0
	a2pid = 0
	t3pid = 0
	a3pid = 0	
	
	if sAction = "" then
    objRS.Open "SELECT * FROM qryPendingTrade WHERE (((qryPendingTrade.aowner)=" & ownerid & ")) order by DecisionDate", objConn,3,3,1
	else
		objRS.Open "SELECT * FROM qryPendingTrade WHERE (((qryPendingTrade.tradeid)=" & tradeid & ")) order by DecisionDate", objConn,3,3,1
	end if
	
	'Response.Write "CHECKPOINT A<br>"
	
	t1first  = objRS.Fields("t1first").Value
	t1last   = objRS.Fields("t1last").Value
	t1pid    = objRS.Fields("t1PID").Value

	if objRS.Fields("t2PID").Value > 0 then
		t2pid  = objRS.Fields("t2PID").Value
		t2first= objRS.Fields("t2first").Value
		t2last = objRS.Fields("t2last").Value
	end if

	if objRS.Fields("t3PID").Value > 0 then
		t3pid  = objRS.Fields("t2PID").Value
		t3first= objRS.Fields("t3first").Value
		t3last = objRS.Fields("t3last").Value
	end if

	a1first  = objRS.Fields("a1first").Value
	a1last   = objRS.Fields("a1last").Value
	a1pid    = objRS.Fields("a1PID").Value

	if objRS.Fields("a2PID").Value > 0 then
		a2pid  = objRS.Fields("a2PID").Value
		a2first= objRS.Fields("a2first").Value
		a2last = objRS.Fields("a2last").Value
	end if

	if objRS.Fields("a3PID").Value > 0 then
		a3pid  = objRS.Fields("a3PID").Value
		a3first= objRS.Fields("a3first").Value
		a3last = objRS.Fields("a3last").Value
	end if

%>
<!--#include virtual="Common/functions.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
</head>
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
    background-color: #d9ded1;
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

<% if sAction = "" then %>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Trade Offers</strong>
			</div>
		</div>
	</div>
</div>
<%
 While Not objRS.EOF
  wTeamName = replace((objRS.Fields("tteamname").Value), "THE ", "")
	
	if len(wTeamName) >19 then 
	 wTeamName = objRS.Fields("tteamnameshort").Value
	end if
%>

<form action="tradeoffers.asp" name="frmMain" method="POST">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("towner").Value %>" />
<input type="hidden" name="var_teamname" value="<%=objRS.Fields("tteamname").Value %>" />
<input type="hidden" name="var_myname" value="<%=objRS.Fields("ateamname").Value %>" />
<input type="hidden" name="var_invtradeind" value="<%=objRS.Fields("InvalidTradeInd").Value %>" />
<input type="hidden" name="var_tradeid" value="<%=objRS.Fields("TradeId").Value %>" />
<input type="hidden" name="var_tradedeleted" value="<%= tradedeleted %>" />
<!--#include virtual="Common/headermain.inc"-->

  <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">	
					<div class="panel panel-override">
            <table class="table table-custom-black  table-bordered table-condensed">
							<tr>
								<th style="width:50%;"><%=wTeamName%></th>
								<th style="width:50%;">My Players</th>
							</tr>
							<tr>
								<td style="width:50%;">
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
								 set playerBarps = 0	 
								 if objrstrade.RecordCount >= 0 then
									playerBarps = objrstrade.Fields("barps").Value
									t1playerBarps = objrstrade.Fields("barps").Value
									t1blks      	= objrstrade.Fields("blk").Value
									t1asts      	= objrstrade.Fields("ast").Value
									t1rebs      	= objrstrade.Fields("reb").Value
									t1pts       	= objrstrade.Fields("ppg").Value 
									t1stls      	= objrstrade.Fields("stl").Value
									t1turns     	= objrstrade.Fields("to").Value
									t1threes    	= objrstrade.Fields("three").Value
									objRSNext5.Open "select gameday from qryMissingLineup " & _
																	"where pid = " & t1pid & " " & _
																	"and gameday >= Date() " & _
																	"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
													 
									objRSPO.Open 		"select gameday from qryMissingLineup " & _
																	"where pid = " & t1pid & " " & _
																	"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
												 
									tplayer1GmCnt  = objRSNext5.RecordCount
									tplayer1POGmCnt= objRSPO.RecordCount
									objRSNext5.close
									objRSPO.close
								 end if

								%>
									
										
											<table class="table table-bordered table-condensed">
													<tr style="background-color:white;"> 
														<td class="big">
															<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t1last").Value,10) %></a>&nbsp;
															<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;|&nbsp;<span ><%=round(t1playerBarps,2) %></span>
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
													set playerBarps = 0	 
													if objRS.Fields("t2PID").Value > 0 then
															objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t2first&"' and lastName = '"&t2last&"' "

														if objrstrade.RecordCount >= 0 then
																t2playerBarps = objrstrade.Fields("barps").Value
																t2blks      = objrstrade.Fields("blk").Value
																t2asts      = objrstrade.Fields("ast").Value
																t2rebs      = objrstrade.Fields("reb").Value
																t2pts       = objrstrade.Fields("ppg").Value 
																t2stls      = objrstrade.Fields("stl").Value
																t2turns     = objrstrade.Fields("to").Value
																t2threes    = objrstrade.Fields("three").Value
																t2pid       = objRS.Fields("t2PID").Value
															playerBarps = objrstrade.Fields("barps").Value
															objRSNext5.Open  	"select gameday from qryMissingLineup " & _
																								"where pid = " & t2pid & " " & _
																								"and gameday >= Date() " & _
																								"and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
																		 
															objRSPO.Open 			"select gameday from qryMissingLineup " & _
																								"where pid = " & t2pid & " " & _
																								"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
												
															tplayer2GmCnt  = objRSNext5.RecordCount
															tplayer2POGmCnt= objRSPO.RecordCount
															objRSNext5.close
															objRSPO.close
														end if
													%>

													<tr style="background-color:white;"> 
														<td class="big">
															<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t2last").Value,10) %></a>&nbsp;
															<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;|&nbsp;<span ><%=round(t2playerBarps,2) %></span>
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
										 set playerBarps = 0	 
										if (objRS.Fields("t3PID").Value) > 0 then
										objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&t3first&"' and lastName = '"&t3last&"' "

											 if objrstrade.RecordCount >= 0 then
												t3playerBarps = objrstrade.Fields("barps").Value
												t3blks      = objrstrade.Fields("blk").Value
												t3asts      = objrstrade.Fields("ast").Value
												t3rebs      = objrstrade.Fields("reb").Value
												t3pts       = objrstrade.Fields("ppg").Value 
												t3stls      = objrstrade.Fields("stl").Value
												t3turns     = objrstrade.Fields("to").Value
												t3threes    = objrstrade.Fields("three").Value
												t3pid       = objRS.Fields("t3PID").Value
												playerBarps = objrstrade.Fields("barps").Value
													
												objRSNext5.Open  "select gameday from qryMissingLineup " & _
																				 "where pid = " & t3pid & " " & _
																				 "and gameday >= Date() " & _
																				 "and gameday < (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1
																 
												objRSPO.Open 		"select gameday from qryMissingLineup " & _
																				"where pid = " & t3pid & " " & _
																				"and gameday >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') ", objConn,3,3,1  
													
												tplayer3GmCnt  = objRSNext5.RecordCount
												tplayer3POGmCnt= objRSPO.RecordCount
												objRSNext5.close
												objRSPO.close
												end if
										%>

										<tr style="background-color:white;"> 
											<td class="big">
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("t3last").Value,10) %></a>&nbsp;
												<span class="gameTip big" "><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;|&nbsp;<span ><%=round(t3playerBarps,2) %></span>
											</td>
										</tr>	
				
										<% end if %>
										</table>
							
								</td>
								<td style="width:50%;">
								 <%
									objrstrade.Close
									objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a1first&"' and lastName = '"&a1last&"' " 
									a1blks   = 0
									a1asts   = 0
									a1rebs   = 0
									a1pts    = 0 
									a1stls   = 0
									a1turns  = 0
									a1threes = 0
									a1playerBarps = 0	 
									set playerBarps = 0	 
									 if objrstrade.RecordCount >= 0 then
										a1playerBarps = objrstrade.Fields("barps").Value
										a1blks      = objrstrade.Fields("blk").Value
										a1asts      = objrstrade.Fields("ast").Value
										a1rebs      = objrstrade.Fields("reb").Value
										a1pts       = objrstrade.Fields("ppg").Value 
										a1stls      = objrstrade.Fields("stl").Value
										a1turns     = objrstrade.Fields("to").Value
										a1threes    = objrstrade.Fields("three").Value		
										playerBarps = objrstrade.Fields("barps").Value
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
									<table class="table table-bordered table-condensed">
										<tr  style="background-color:white">
											<td class="big">
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a1last").Value,10) %></a>&nbsp;
												<span class="gameTip big" "><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;|&nbsp;<span ><%=round(a1playerBarps,2) %></span>
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
										set playerBarps = 0	
										if (objRS.Fields("a2PID").Value) > 0 then
										 objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a2first&"' and lastName = '"&a2last&"' "
			 
											 if objrstrade.RecordCount >= 0 then
												a2playerBarps = objrstrade.Fields("barps").Value
												a2blks      = objrstrade.Fields("blk").Value
												a2asts      = objrstrade.Fields("ast").Value
												a2rebs      = objrstrade.Fields("reb").Value
												a2pts       = objrstrade.Fields("ppg").Value 
												a2stls      = objrstrade.Fields("stl").Value
												a2turns     = objrstrade.Fields("to").Value
												a2threes    = objrstrade.Fields("three").Value
												a2pid       = objRS.Fields("a2PID").Value
															 
												playerBarps = objrstrade.Fields("barps").Value
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
											<td class="big">
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a2last").Value,10) %></a>&nbsp;
												<span class="gameTip big" "><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;|&nbsp;<span ><%=round(a2playerBarps,2) %></span>
											</td>
										</tr>	
		
										<% end if %>		
										<% 
										objrstrade.Close
										'Response.Write "t2PID = "&t2PID&" <br>"
										a3blks   = 0
										a3asts   = 0
										a3rebs   = 0
										a3pts    = 0 
										a3stls   = 0
										a3turns  = 0
										a3threes = 0	
										a3playerBarps = 0	
										set playerBarps = 0	 
										if (objRS.Fields("a3PID").Value) > 0 then
										objrstrade.Open "SELECT * FROM qry_tblbarps WHERE firstName = '"&a3first&"' and lastName = '"&a3last&"' " 

											if objrstrade.RecordCount >= 0 then
												a3playerBarps = objrstrade.Fields("barps").Value
												a3blks        = objrstrade.Fields("blk").Value
												a3asts        = objrstrade.Fields("ast").Value
												a3rebs        = objrstrade.Fields("reb").Value
												a3pts         = objrstrade.Fields("ppg").Value 
												a3stls        = objrstrade.Fields("stl").Value
												a3turns       = objrstrade.Fields("to").Value
												a3threes      = objrstrade.Fields("three").Value	
												a3pid         = objRS.Fields("a3PID").Value
												playerBarps = objrstrade.Fields("barps").Value
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
											<td class="big">
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>"><%=left(objRS.Fields("a3last").Value,10) %></a>&nbsp;
												<span class="gameTip big" "><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;|&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>&nbsp;|&nbsp;<span ><%=round(a3playerBarps,2) %></span>
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
						diffBarps   = cDbl(tbarps) -  cDbl(abarps)
						
						tblks       = cDbl(t1blks) + cDbl(t2blks) + cDbl(t3blks)
						ablks       = cDbl(a1blks) + cDbl(a2blks) + cDbl(a3blks)
						diffblks    = cDbl(tblks) - cDbl(ablks)
						
						tasts       = cDbl(t1asts) + cDbl(t2asts) + cDbl(t3asts)
						aasts       = cDbl(a1asts) + cDbl(a2asts) + cDbl(a3asts)
						diffasts    = cDbl(tasts) - cDbl(aasts)

						trebs       = cDbl(t1rebs) + cDbl(t2rebs) + cDbl(t3rebs)
						arebs       = cDbl(a1rebs) + cDbl(a2rebs) + cDbl(a3rebs)
						diffrebs    = cDbl(trebs) - cDbl(arebs)								

						tpts        = cDbl(t1pts) + cDbl(t2pts) + cDbl(t3pts)
						apts        = cDbl(a1pts) + cDbl(a2pts) + cDbl(a3pts)
						diffpts     = cDbl(tpts) - cDbl(apts)

						tstls       = cDbl(t1stls) + cDbl(t2stls) + cDbl(t3stls)
						astls       = cDbl(a1stls) + cDbl(a2stls) + cDbl(a3stls)
						diffstls    = cDbl(tstls) - cDbl(astls)

						ttos        = cDbl(t1turns) + cDbl(t2turns) + cDbl(t3turns)
						atos        = cDbl(a1turns) + cDbl(a2turns) + cDbl(a3turns)
						'May have to flip
						difftos     = cDbl(ttos) - cDbl(atos)

						tthrees     = cDbl(t1threes) + cDbl(t2threes) + cDbl(t3threes)
						athrees     = cDbl(a1threes) + cDbl(a2threes) + cDbl(a3threes)
						diffthrees  = cDbl(tthrees) - cDbl(athrees)
				
						%>
					<table class="table table-custom-black  table-bordered table-condensed">
					<tr>
						<td class="big" style="background-color:#212121;font-size:14px;color:#FFEB3B;font-weight:700;text-align: center;" colspan="8">Plus/Minus By Stats</td>
					</tr>
					<tr>
						<th class="big"  width="14%"><span style="color:black;">B</span>/pg</th>
						<th class="big"  width="14%"><span style="color:black;">A</span>/pg</th>
						<th class="big"  width="14%"><span style="color:black;">R</span>/pg</th>
						<th class="big"  width="14%"><span style="color:black;">P</span>/pg</th>
						<th class="big"  width="14%"><span style="color:black;">S</span>/pg</th>
						<th class="big"  width="14%"><span style="color:black;">3</span>/pg</th>
						<th class="big"  width="16%"><span style="color:black;">T</span>/pg</th>	
					</tr>	
					<tr style="background-color:#212121;font-size:14px;color:#FFEB3B;font-weight:700;">
						<% if (cDbl(tblks)) > (cDbl(ablks)) then %>
								<td class="big"  align="center">+<%=round(diffblks,1)%></td>
						<% elseif (cDbl(tblks)) <(cDbl(ablks)) then%>
								<td class="big"  align="center"><%=round(diffblks,1)%></td>
						<% elseif (cDbl(ablks)) = (cDbl(tblks)) then %>
								<td class="big" align="center">E</td>
						<% end if %>
						
						<% if tasts > aasts then %>
								<td class="big" align="center">+<%=round(diffasts,1)%></span></td>
						<% elseif tasts < aasts then %>
								<td class="big" align="center"><%=round(diffasts,1)%></span></td>
						<% elseif aasts = tasts then %>
								<td class="big" align="center">E</span></td>
						<% end if %>
						
						<% if trebs > arebs then %>
								<td class="big" align="center">+<%=round(diffrebs,1)%></td>
						<% elseif trebs < arebs then %>
								<td class="big"  align="center"><%=round(diffrebs,1)%></td>
						<% elseif trebs = arebs then %>
								<td class="big" align="center">E</td>
						<% end if %>
						
						<% if (cDbl(tpts)) > (cDbl(apts)) then %>
								<td class="big" align="center">+<%=round(diffpts,1)%></td>
						<% elseif (cDbl(tpts)) < (cDbl(apts)) then %>
								<td class="big" align="center"><%=round(diffpts,1)%></td>
						<% elseif (cDbl(tpts)) = (cDbl(apts))then %>
								<td class="big" align="center">E</td>
						<% end if %>
						<% if tstls > astls then %>
								<td class="big" align="center">+<%=round(diffstls,1)%></td>
						<% elseif tstls < astls then %>
								<td class="big" align="center"><%=round(diffstls,1)%></td>
						<% elseif tstls = astls then %>
								<td class="big" align="center">E</td>
						<% end if %>

						<% if tthrees > athrees then %>
								<td class="big" align="center">+<%=round(diffthrees,1)%></td>
						<% elseif tthrees < athrees then %>
								<td class="big" align="center"><%=round(diffthrees,1)%></td>
						<% elseif tthrees = athrees then %>
								<td class="big" align="center">E</td>
						<% end if %>
						
						<% if ttos > atos then %>
								<td class="big" align="center"><%=round(difftos,1)%></td>
						<% elseif ttos < atos then%>
								<td class="big" align="center">+<%=round(difftos*-1,1)%></td>
						<% elseif ttos = atos then%>		
								<td class="big" align="center">E</td>
						<% end if %>
						</tr>
						<tr>
						<td class="big" colspan="7" style="background-color:#FFEB3B;" align="center"><span style="font-weight:bold;color:black;" class="text-uppercase"><i class="fa fa-exclamation-triangle" aria-hidden="true"></i>&nbsp;Net result of this Trade is <span> 
						<% if tbarps > abarps then %>
								<span class="badgeUp" data-toggle="tooltip" title="Barps Diff">+<%=round(diffBarps,1)%></span> Barps!
						<% elseif tbarps < abarps then  %>
								<span class="badgeDown" data-toggle="tooltip" title="Barps Diff"><%=round(diffBarps,1)%></span> Barps!
						<% elseif tbarps = abarps then %>
								<span class="badgeEven" data-toggle="tooltip" title="Barps Diff">E</span></span> Barps!
						<% end if %>
						</td>
						</tr>
						<tr style="background-color:black;font-weight:bold;color:yellowgreen;">
							<td style="text-align:center;" colspan="7">This Trade offer expires around <%=objRS.Fields("DecisionDate").Value%></td>
						</tr>
						<tr>
							<td class="big"  colspan="8"><textarea name="txtNotes" class="form-control" rows="3" placeholder="Enter Trade Comments" id="txtNotes"></textarea>
						</tr>

							<tr>
							<td class="big" colspan="8">
									<button type="submit" value="Accept"   name="Action" class="btn btn-default-green btn-block btn-md"><i class="fa fa-handshake-o" aria-hidden="true"></i>&nbsp;Accept</button>
									<button type="submit" value="Forecast" name="Action" class="btn btn-default-blue btn-block btn-md"><i class="fa fa-bar-chart" aria-hidden="true"></i>&nbsp;Forecast</button>
									<button type="submit" value="Counter"  name="Action" class="btn btn-default btn-block btn-md"><i class="fa fa-exchange" aria-hidden="true"></i>&nbsp;Counter</button>
									<button type="submit" value="Decline"  name="Action" class="btn btn-default-red btn-block btn-md"><i class="fas fa-trash-alt"></i>&nbsp;Decline</button>

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
End if
if sAction = "Accept" and errorCode = "Confirmation" then %>


<%
 While Not objRS.EOF
	wTeamName = replace((objRS.Fields("tteamname").Value), "THE ", "")
	if len(wTeamName) >19 then 
	 wTeamName = objRS.Fields("tteamnameshort").Value
	end if
%>
<form action="tradeoffers.asp" name="frmMain" method="POST">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("towner").Value %>" />
<input type="hidden" name="var_teamname" value="<%=objRS.Fields("tteamname").Value %>" />
<input type="hidden" name="var_myname" value="<%=objRS.Fields("ateamname").Value %>" />
<input type="hidden" name="var_invtradeind" value="<%=objRS.Fields("InvalidTradeInd").Value %>" />
<input type="hidden" name="var_tradeid" value="<%=objRS.Fields("TradeId").Value %>" />
<input type="hidden" name="var_tradedeleted" value="<%= tradedeleted %>" />
<!--#include virtual="Common/headermain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Trade Confirmation</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">	
				<div class="panel panel-override">
					<table class="table table-custom-black  table-responsive table-bordered table-condensed">
					<tr>
						<th style="width:50%;text-align:left;"><%=wTeamName%></th>
						<th style="width:50%;text-align:left;">My Players</th>
					</tr>
					<%
						objrstrade.Open "SELECT * FROM qry_tblbarps WHERE first = '"&t1first&"' and last = '"&t1last&"' "
					%>
					<tr>
						<td>
									<table class="table table-bordered table-condensed">
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("t1first").Value) + len(objRS.Fields("t1last").Value)) >= 16 then %>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t1first").Value,1)%>.&nbsp;
													<%=left(objRS.Fields("t1last").Value,14)%>
												</a>,&nbsp;
											<%else%>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t1first").Value,8)%>&nbsp
													<%=left(objRS.Fields("t1last").Value,10)%>
												</a>,&nbsp;
											<%end if%>
												<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span></small
											</td>
										</tr>
									<%
									objrstrade.Close
									if objRS.Fields("t2PID").Value > 0 then
										objrstrade.Open "SELECT * FROM qry_tblbarps WHERE first = '"&t2first&"' and last = '"&t2last&"' "
									%>
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("t2first").Value) + len(objRS.Fields("t2last").Value)) >= 16 then %>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t2first").Value,1)%>.&nbsp;
													<%=left(objRS.Fields("t2last").Value,14)%>
												</a>,&nbsp;
											<%else%>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t2first").Value,8)%>&nbsp;
													<%=left(objRS.Fields("t2last").Value,10)%>
												</a>,&nbsp;
											<%end if%>
												<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>
											</td>
										</tr>
									<%end if%>
									              <%
									objrstrade.Close
									if (objRS.Fields("t3PID").Value) > 0 then
										objrstrade.Open "SELECT * FROM qry_tblbarps WHERE first = '"&t3first&"' and last = '"&t3last&"' "
									%>
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("t3first").Value) + len(objRS.Fields("t3last").Value)) >= 16 then %>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t3first").Value,1)%>.&nbsp;
													<%=left(objRS.Fields("t3last").Value,14)%>
												</a>,&nbsp;
											<%else%>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("t3first").Value,8)%>&nbsp; 
													<%=left(objRS.Fields("t3last").Value,10)%>
												</a>,&nbsp;
											<%end if%>
												<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>
											</td>
										</tr>
									<% end if %>
							</table>							
						</td>
						<td>
								<table class="table table-bordered table-condensed">
								<%
								objrstrade.Close
								objrstrade.Open "SELECT * FROM qry_tblbarps WHERE first = '"&a1first&"' and last = '"&a1last&"' " 
								%>
										<tr style="background-color:white">
											<td>
											<%if (len(objRS.Fields("a1first").Value) + len(objRS.Fields("a1last").Value)) >= 16 then %>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("a1first").Value,1)%>.&nbsp;
													<%=left(objRS.Fields("a1last").Value,14)%>
												</a>,&nbsp;
											<%else%>
												<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
													<%=left(objRS.Fields("a1first").Value,8)%>&nbsp;
													<%=left(objRS.Fields("a1last").Value,10)%>
												</a>,&nbsp;
											<%end if%>
												<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>
											</td>
										</tr>
								<%
								objrstrade.Close
								if (objRS.Fields("a2PID").Value) > 0 then
								objrstrade.Open "SELECT * FROM qry_tblbarps WHERE first = '"&a2first&"' and last = '"&a2last&"' " %>
									<tr style="background-color:white">
										<td>
										<%if (len(objRS.Fields("a2first").Value) + len(objRS.Fields("a2last").Value)) >= 16 then %>
											<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
												<%=left(objRS.Fields("a2first").Value,1)%>.&nbsp;
												<%=left(objRS.Fields("a2last").Value,14)%>
											</a>,&nbsp;
										<%else%>
											<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
												<%=left(objRS.Fields("a2first").Value,8)%>&nbsp;
												<%=left(objRS.Fields("a2last").Value,10)%>
											</a>,&nbsp;
										<%end if%>
											<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>
										</td>
									</tr>
								<% end if %>	
								<% 
								objrstrade.Close
								if (objRS.Fields("a3PID").Value) > 0 then
								objrstrade.Open "SELECT * FROM qry_tblbarps WHERE first = '"&a3first&"' and last = '"&a3last&"' " 
								%>
									<tr style="background-color:white">
										<td>
										<%if (len(objRS.Fields("a3first").Value) + len(objRS.Fields("a3last").Value)) >= 16 then %>
											<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
												<%=left(objRS.Fields("a3first").Value,1)%>.&nbsp;
												<%=left(objRS.Fields("a3last").Value,14)%>
											</a>,&nbsp;
										<%else%>
											<a class="blue" href="playerprofile.asp?pid=<%=objrstrade.Fields("PID").Value %>">
												<%=left(objRS.Fields("a3first").Value,8)%>&nbsp;
												<%=left(objRS.Fields("a3last").Value,10)%>
											</a>,&nbsp;
										<%end if%>
											<span class="gameTip big"><%= objrstrade.Fields("teamShortName").Value %></span>&nbsp;<span class="big orangeText"><%=objrstrade.Fields("pos").Value %></span>
										</td>
									</tr>	
								<% end if %>
							</table>							
						</td>
					</tr>
					<%
									objrstrade.Close 
							%>
								<tr>
									<td colspan="6"><textarea name="txtNotes" class="form-control" rows="3" placeholder="Comments" id="txtNotes"></textarea>
								</tr>
								<tr>
									<td colspan="6">
										<button type="submit" value="Trade Confirmation" name="Action" class="btn btn-block btn-default-green btn-md"><span class="glyphicon glyphicon-save"></span>&nbsp;Confirm</button>
										<button type="submit" value="Decline" name="Action" class="btn btn-default-red  btn-block btn-md"><i class="fas fa-trash-alt"></i>&nbsp;Decline</button>
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
 	t1first  = objRS.Fields("t1first").Value
	t1last   = objRS.Fields("t1last").Value

	if objRS.Fields("t2PID").Value > 0 then
		t2first   = objRS.Fields("t2first").Value
		t2last    = objRS.Fields("t2last").Value
	end if

	if objRS.Fields("t3PID").Value > 0 then
		t3first  = objRS.Fields("t3first").Value
  		t3last   = objRS.Fields("t3last").Value
	end if

	a1first  = objRS.Fields("a1first").Value
	a1last   = objRS.Fields("a1last").Value

	if objRS.Fields("a2PID").Value > 0 then
		a2first   = objRS.Fields("a2first").Value
		a2last   = objRS.Fields("a2last").Value
	end if

	if objRS.Fields("a3PID").Value > 0 then
		a3first  = objRS.Fields("a3first").Value
		a3last   = objRS.Fields("a3last").Value
	end if

	Wend

%>
<%
End if
if sAction = "Accept" and errorcode = "Trade Deleted"  then %>
<form action="tradeoffers.asp" name="frmDelete" method="POST">
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("towner").Value %>" />
  <input type="hidden" name="var_teamname" value="<%=objRS.Fields("tteamname").Value %>" />
  <input type="hidden" name="var_myname" value="<%=objRS.Fields("ateamname").Value %>" />
  <input type="hidden" name="var_tradedeleted" value="<%= tradedeleted %>" />
  <!--#include virtual="Common/headermain.inc"-->
  <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">
        <div class="panel panel-danger">
          <div class="panel-heading clearfix"> <i class="icon-calendar"></i>
            <h3 class="panel-title">Trade Request Error</h3>
          </div>
          <div class="panel-body">
            <table class="table table-condensed">
              <tr>
                <td>The trade you specified cannot be processed due to an invalid browser 
                  state. Please refresh your pending trades and select the trade again. 
                  Reasons this can happen:<br>
                  <ul type="bullet">
                    <li>You hit the browser&#39;s back button after deleting the trade offer, 
                      so that the trade you are trying to accept is no longer there. </li>
                    <li>You are using multiple browser windows and are trying to access 
                      a trade that has been deleted.</li>
                  </ul></td>
              </tr>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
</form>
<%
End if
 'Response.Write "errorCode = "&errorCode&"<br>"
if sAction = "Accept" and errorcode = "Trade Deadline Passed" then %>
<form action="tradeoffers.asp" name="frmtradeviolation" method="POST">
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("towner").Value %>" />
  <input type="hidden" name="var_teamname" value="<%=objRS.Fields("tteamname").Value %>" />
  <input type="hidden" name="var_myname" value="<%=objRS.Fields("ateamname").Value %>" />
  <input type="hidden" name="var_tradedeleted" value="<%= tradedeleted %>" />
  <!--#include virtual="Common/headermain.inc"-->
  <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">
        <div class="panel panel-danger">
          <div class="panel-heading clearfix"> <i class="icon-calendar"></i>
            <h3 class="panel-title">Trade Deadline Passed</h3>
          </div>
          <div class="panel-body">
            <table class="table table-condensed">
              <tr>
                <td>Too Late.  Trade Deadline has passed! </td>
              </tr>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
</form>
<%
End if
if sAction = "Trade Confirmation" and errorcode = "Trade Violation" then %>
<form action="tradeoffers.asp" name="frmtradeviolation" method="POST">
  <input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
  <input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("towner").Value %>" />
  <input type="hidden" name="var_teamname" value="<%=objRS.Fields("tteamname").Value %>" />
  <input type="hidden" name="var_myname" value="<%=objRS.Fields("ateamname").Value %>" />
  <input type="hidden" name="var_tradedeleted" value="<%= tradedeleted %>" />
  <!--#include virtual="Common/headermain.inc"-->
  <div class="container">
    <div class="row">
      <div class="col-md-12 col-sm-12 col-xs-12">
        <div class="panel panel-danger">
          <div class="panel-heading clearfix"> <i class="icon-calendar"></i>
            <h3 class="panel-title">Trade Request Error</h3>
          </div>
          <div class="panel-body">
            <table class="table table-condensed">
              <tr>
                <td>Processing this trade violates the legal Roster Limit of 14! </td>
              </tr>
            </table>
          </div>
        </div>
      </div>
    </div>
  </div>
</form>
<%
End if
if sAction = "Trade Confirmation" and errorcode = "In Play Violation" then %>
<form action="tradeoffers.asp" name="frmtradeviolation" method="POST">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_tradepartner" value="<%=objRS.Fields("towner").Value %>" />
<input type="hidden" name="var_teamname" value="<%=objRS.Fields("tteamname").Value %>" />
<input type="hidden" name="var_myname" value="<%=objRS.Fields("ateamname").Value %>" />
<input type="hidden" name="var_tradedeleted" value="<%= tradedeleted %>" />
<!--#include virtual="Common/headermain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
				<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> Trading Players Active!<br>
				<%=errormessage %>			
			</div>
		</div>
	</div>
</div>
</form>
<%
End if
if sAction = "Forecast" or sAction = "Reforecast" then
		
   tradeid        = Request.Form("var_tradeid")
	 'Response.Write " trade id value  = "&tradeid&" <br> "
	 
	 if sAction = "Forecast" then 
		wNeutralVal = false
	 else
		wNeutralVal = true
	 end if	

	
   objrstrade.Open "SELECT t.* FROM tblPendingTrades t WHERE t.tradeid = " & tradeid, objConn,3,3,1						
   tplayer1       = objrstrade.Fields("TradedPlayerID").Value
   tplayer2       = objrstrade.Fields("TradedPlayerID2").Value
   tplayer3       = objrstrade.Fields("TradedPlayerID3").Value
   aplayer1       = objrstrade.Fields("AcquiredPlayerID").Value
   aplayer2       = objrstrade.Fields("AcquiredPlayerID2").Value
   aplayer3       = objrstrade.Fields("AcquiredPlayerID3").Value
   TradeOwnerId   = objrstrade.Fields("FromOID").Value
   objrstrade.Close
   
   if tplayer1 <> 0 then FuncCall = Get_One_Name(tplayer1, stPlayer1Name, x, stRegGameCnt1, stPOGameCnt1) else stPlayer1Name = "" end if
   if tplayer2 <> 0 then FuncCall = Get_One_Name(tplayer2, stPlayer2Name, x, stRegGameCnt2, stPOGameCnt2) else stPlayer2Name = "" end if
   if tplayer3 <> 0 then FuncCall = Get_One_Name(tplayer3, stPlayer3Name, x, stRegGameCnt3, stPOGameCnt3) else stPlayer3Name = "" end if
   if aplayer1 <> 0 then FuncCall = Get_One_Name(aplayer1, saPlayer1Name, x, saRegGameCnt1, saPOGameCnt1) else saPlayer1Name = "" end if
   if aplayer2 <> 0 then FuncCall = Get_One_Name(aplayer2, saPlayer2Name, x, saRegGameCnt2, saPOGameCnt2) else saPlayer2Name = "" end if
   if aplayer3 <> 0 then FuncCall = Get_One_Name(aplayer3, saPlayer3Name, x, saRegGameCnt3, saPOGameCnt3) else saPlayer3Name = "" end if
   
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
      gameday       = objRSAll("gameday")
	
    wRetcd = Forecast_Lineup(gameday,ownerid,0,CenName,CenPID,CenBarps,For1Name,For1PID,For1Barps,For2Name,For2Pid,For2Barps,Gua1Name,Gua1PID,Guard1Barps,Gua2Name,Gua2PID,Guard2Barps)
	  wRetcd = Forecast_Lineup(gameday,ownerid,TradeOwnerId,OppCenName,OppCenPID,OppCenBarps,OppFor1Name,OppFor1PID,OppFor1Barps,OppFor2Name,OppFor2PID,OppFor2Barps,OppGua1Name,OppGua1PID,OppGuard1Barps,OppGua2Name,OppGua2PID,OppGuard2Barps)	

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
	        	  
      objRSAll.MoveNext
   Wend
   objRSAll.Close	
   
   'Response.Write "FC wPositive= "&wPositive&", wNegative= "&wNegative&", wEven="&wEven&"<br>"
	
%>
<form action="tradeoffers.asp" name="frmtradeviolation" method="POST" onSubmit="return processReforecast(this)">
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
	
	wRetcd = Forecast_Lineup(gameday,ownerid,0,CenName,CenPID,CenBarps,For1Name,For1PID,For1Barps,For2Name,For2Pid,For2Barps,Gua1Name,Gua1PID,Guard1Barps,Gua2Name,Gua2PID,Guard2Barps)
	wRetcd = Forecast_Lineup(gameday,ownerid,TradeOwnerId,OppCenName,OppCenPID,OppCenBarps,OppFor1Name,OppFor1PID,OppFor1Barps,OppFor2Name,OppFor2PID,OppFor2Barps,OppGua1Name,OppGua1PID,OppGuard1Barps,OppGua2Name,OppGua2PID,OppGuard2Barps)	
	
	myTotal  = cDbl(CenBarps)+cDbl(For1Barps)+cDbl(For2Barps)+cDbl(Guard1Barps)+cDbl(Guard2Barps)
	OppTotal = cDbl(OppCenBarps)+cDbl(OppFor1Barps)+cDbl(OppFor2Barps)+cDbl(OppGuard1Barps)+cDbl(OppGuard2Barps)
	xdiff = MyTotal - OppTotal
	if xdiff < 0 then
	     xdiff = xdiff * -1
	end if
	
	'MyFav    = (myTotal - OppTotal)
	'OppFav   = (OppTotal - myTotal)
	
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
		   OpponentName = objRSOpponentName("HomeTeamShort")
		end if	
	else
		OpponentName = "TBD"
	end if					
	
	objRSOpponentName.Close	
%>

<!-- Trade Forecaster Logic-->
					<table class="table table-custom-black table-bordered table-condensed table-responsive">
						<tr> 
							<th style="font-size:12px !important;"><%=(FormatDateTime(objRSAll("gameday"),1))%></th>
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
					<% if (CenPID = aplayer1 or CenPID = aplayer2 or CenPID = aplayer3) and CenPID <> 0  then %>
							<td style="text-align: center"><blueText><%=CenName%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=CenName%></td>
					<% end if %>		
						<td class="big " style="text-align: center;"><%=round(CenBarps,2)%></td>
						<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">CEN</th>		
						<td class="big " style="text-align: center;"><%=round(OppCenBarps,2)%></td>					
					<% if (OppCenPID = tplayer1 or OppCenPID = tplayer2 or OppCenPID = tplayer3) and OppCenPID <> 0  then %>
							<td style="text-align: center"><blueText><%=OppCenName%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=OppCenName%></td>
					<% end if %>
					</tr>  

					<tr bgcolor="#FFFFFF">
					<% if (For1PID = aplayer1 or For1PID = aplayer2 or For1PID = aplayer3) and For1PID <> 0 then %>
							<td style="text-align: center"><blueText><%=For1Name%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=For1Name%></td>
					<% end if %>	
					<td class="big " style="text-align: center"><%=round(For1Barps,2)%></td>		
					<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">FOR</th>							
					<td class="big " style="text-align: center"><%=round(OppFor1Barps,2)%></td>
					<% if (OppFor1PID = tplayer1 or OppFor1PID = tplayer2 or OppFor1PID = tplayer3) and OppFor1PID <> 0 then %>
						<td style="text-align: center"><blueText><%=OppFor1Name%></blueText></td>
					<% else %>
						<td style="text-align: center"><%=OppFor1Name%></td>
					<% end if %>
					</tr>  
					
					<tr bgcolor="#FFFFFF">
					<% if (For2PID = aplayer1 or For2PID = aplayer2 or For2PID = aplayer3) and For2PID <> 0   then %>
							<td style="text-align: center"><blueText><%=For2Name%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=For2Name%></td>
					<% end if %>						
					<td class="big " width="10%" style="text-align: center"><%=round(For2Barps,2)%></td>
					<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">FOR</th>	
					<td class="big " width="10%" style="text-align: center"><%=round(OppFor2Barps,2)%></td>
					<% if (OppFor2PID = tplayer1 or OppFor2PID = tplayer2 or OppFor2PID = tplayer3) and OppFor2PID <> 0   then %>
							<td style="text-align: center"><blueText><%=OppFor2Name%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=OppFor2Name%></td>
					<% end if %>						
					</tr>  
					
					<tr bgcolor="#FFFFFF">
					<% if (Gua1PID = aplayer1 or Gua1PID = aplayer2 or Gua1PID = aplayer3) and Gua1PID <> 0  then %>
							<td style="text-align: center"><blueText><%=Gua1Name%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=Gua1Name%></td>
					<% end if %>					
					<td class="big " width="10%" style="text-align: center"><%=round(Guard1Barps,2)%></td>
					<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">GUA</th>	
					<td class="big " width="10%" style="text-align: center"><%=round(OppGuard1Barps,2)%></td>
					<% if (OppGua1PID = tplayer1 or OppGua1PID = tplayer2 or OppGua1PID = tplayer3) and OppGua1PID <> 0  then %>
							<td style="text-align: center"><blueText><%=OppGua1Name%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=OppGua1Name%></td>
					<% end if %>
					</tr>  
					
					<tr bgcolor="#FFFFFF">
					<% if (Gua2PID = aplayer1 or Gua2PID = aplayer2 or Gua2PID = aplayer3) and Gua2PID <> 0 then %>
							<td style="text-align: center"><blueText><%=Gua2Name%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=Gua2Name%></td>
					<% end if %>				
					<td class="big " width="10%" style="text-align: center"><%=round(Guard2Barps,2)%></td>
					<th class="text-center big orangeText" style="vertical-align:middle;width:10%;">GUA</th>	
					<td class="big " width="10%" style="text-align: center"><%=round(OppGuard2Barps,2)%></td>
					<% if (OppGua2PID = tplayer1 or OppGua2PID = tplayer2 or OppGua2PID = tplayer3) and OppGua2PID <> 0 then %>
							<td style="text-align: center"><blueText><%=OppGua2Name%></blueText></td>
					<% else %>
							<td style="text-align: center"><%=OppGua2Name%></td>
					<% end if %>

					</tr> 
					<tr style="background-color:white;text-align:center;vertical-align:middle;font-weight: bold;">
						<td class="big text-right">Totals&nbsp;<span style="text-align: right;"><i class="fal fa-arrow-to-right red"></i></span></td>
						<td class="big"><%= round(myTotal,2)%></td>
						<td class="big" style="vertical-align: middle";>
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
				<a HREF="tradeoffers.asp" onClick="history.back();return false;">
					<button type="submit" value="Cancel" formnovalidate  name="Cancel" class="btn btn-default  btn-md"><span class="glyphicon glyphicon-backward"></span>&nbsp;Cancel</button>
				</a> 
		</div>

	</div>
</div>
		<br>
</form>
<%
  end if
  objRS.Close
  Set objRS = Nothing
  ObjConn.Close
  Set objConn = Nothing
  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>
