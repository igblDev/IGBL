<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1302
	Err.Clear
	strErrorUrl = ""

	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"



	'**********************************************************************************************'
	'**** Lineup Code Beginning
	'**** Set Objects in preparation to display the screen
	'**********************************************************************************************'

	Dim objConn, sTeam, sAction, sURL,ownerid,errorCode
	Dim int_TotalBARPS, centerCnt,forwardCnt,guardCnt,unAvailCnt,objRSNext5,objRSteamsMyIGBL
	Dim sCenterGameTime,sForwardGameTime,sForward2GameTime,sGuardGameTime,sGuard2GameTime,bCenterGameTime,bForwardGameTime,bForward2GameTime,bGuardGameTime,			bGuard2GameTime,bkCenter,bkForward1,bkForward2,bkGuard1,bkGuard2,pCenter
	pForward,pGuard,objRSHome, objRSAway, objRSAll 
	Dim lineupFnd,objRSLU,cDeadlinePassed,f1DeadlinePassed,f2DeadlinePassed,g1DeadlinePassed,g2DeadlinePassed
	Dim objRS,objRSUnavail,objRSWork,objRSall,objRSteams,objRSflex, objRSwaivers,objRSgames,objRS1
	Dim objRSGameStarted,objRStradesmade,objRStradesreceived,objRSgameDate,objRSPLineups,selectDate,currentDate,objRSteamRec

	Set objRS              = Server.CreateObject("ADODB.RecordSet")
	Set objRSteamsMyIGBL   = Server.CreateObject("ADODB.RecordSet")
	Set objRSUnavail       = Server.CreateObject("ADODB.RecordSet")
	Set objRSGameStarted   = Server.CreateObject("ADODB.RecordSet")
	Set objRSteams         = Server.CreateObject("ADODB.RecordSet")
	Set objRSgames         = Server.CreateObject("ADODB.RecordSet")
	Set objRSgameDate      = Server.CreateObject("ADODB.RecordSet")
	Set objRSteamRec       = Server.CreateObject("ADODB.RecordSet")
	Set objRSLU            = Server.CreateObject("ADODB.RecordSet")
	Set objRSNext5         = Server.CreateObject("ADODB.RecordSet")
	Set objRSwaivers       = Server.CreateObject("ADODB.RecordSet")
	Set objRStradesreceived= Server.CreateObject("ADODB.RecordSet")
	Set objRStradesmade    = Server.CreateObject("ADODB.RecordSet")
	Set objrsMoney         = Server.CreateObject("ADODB.RecordSet")
	Set objrsForecast      = Server.CreateObject("ADODB.RecordSet")
	Set objRSActive        = Server.CreateObject("ADODB.RecordSet") 
	Set objRSWork          = Server.CreateObject("ADODB.RecordSet") 
	Set objsrCycle         = Server.CreateObject("ADODB.RecordSet")  
	Set objsrBonus         = Server.CreateObject("ADODB.RecordSet")  
	Set objsrAvailPlayers  = Server.CreateObject("ADODB.RecordSet")  
	
	GetAnyParameter "cmbTeam", sTeam
	
	sAction    = Request.querystring("Action")	
	
	if sAction = "" then
		GetAnyParameter "Action", sAction
	end if
	
	'**********************************************************************************************'
	'**** End of Setting Objects for Lineup Display
	'***********************************************************************************************'
%>
<!--#include virtual="Common/session.inc"-->
<%
	objRS.Open                  "SELECT * FROM qry_PlayerAll WHERE (((qry_PlayerAll.OwnerID)=" & ownerid & "))", objConn
	objRSteamsMyIGBL.Open       "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & ownerid & "))", objConn
	objRSwaivers.Open	        	"SELECT * FROM tblWaivers WHERE (((tblWaivers.OwnerID)=" & ownerid & "))", objConn,3,3,1
	objRStradesmade.Open	    	"SELECT * FROM tblPendingTrades WHERE (((tblPendingTrades.fromoid)=" & ownerid & "))", objConn,3,3,1
	objRStradesreceived.Open		"SELECT * FROM tblPendingTrades WHERE (((tblPendingTrades.tooid)=" & ownerid & "))", objConn,3,3,1
	objrsForecast.Open		      "SELECT * FROM tblTradeAnalysis WHERE (((tblTradeAnalysis.fromoid)=" & ownerid & "))", objConn,3,3,1
 	objrsMoney.Open             "SELECT * FROM qrymoney where ownerid = " & ownerid & " ", objConn
	objsrCycle.Open             "SELECT max (cycle) as currentCycle from Standings_Cycle", objConn,3,3,1
		
	w_current_cycle = objsrCycle.Fields("currentCycle").Value
	objsrCycle.Close
	
	
	'Response.Write "MAX CYCLE      = "&w_current_cycle&".<br>"
	
	if ISNULL(w_current_cycle) then
		cycleWins       = 0
	  cycleLoss       = 0
	  w_current_cycle = 0
	else
		objsrBonus.open  "SELECT * From Standings_Cycle WHERE cycle = "& w_current_cycle &" and ID = "& ownerid &" ",objConn,3,3,1 											
		cycleWins = objsrBonus.Fields("won").Value
		cycleLoss = objsrBonus.Fields("loss").Value
		'Response.Write "WINS      = "&cycleWins&".<br>"
		'Response.Write "LOSS      = "&cycleLoss&".<br>"	
	   objsrBonus.Close
	end if
		
	Logo = objRSteamsMyIGBL.Fields("TeamLogo").Value
	
	w_total_spent = objrsMoney.Fields("TotalSpent").Value
	w_tm_rank = objRSteamsMyIGBL.Fields("rank").value

	objrsMoney.Close
	set objrsMoney = nothing
	if 	w_total_spent <= 0 then
		w_total_spent = 0
	end if
	ownerid = CInt(ownerid)  'Convert the variant to a Integer	

	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE_DEADLINE' ",objConn
	wTradeDeadLine = objRSWork.Fields("param_date").Value
	objRSWork.Close
	
	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PLAYOFF_START_DATE' ",objConn
	wPlayoffStart = objRSWork.Fields("param_date").Value
	objRSWork.Close
	
	'**********************************************************************************************'
	'**** Evaluate Action Fired by the User
	'***********************************************************************************************'
	select case sAction
	  case "Submit Lineup"
		errorCode = ""
		errorCodeLU = ""

		Dim strSQL,sCenter,sForward,sGuard2,sForward2,sGuard,sDate,gamedate,barps
		Dim centerCBoxCnt,fowardCBoxCnt,guardCBoxCnt,int_TotalCboxes,foward2CBoxCnt,guard2CBoxCnt  
		
		sDate    	  	   = Now()
		ownerid     	   = Request.Form("var_ownerid")
		gamedate         = Request.Form("GameDays")
		centerCnt     	 = Request.Form("centerCnt")
		forwardCnt     	 = Request.Form("forwardCnt")
		guardCnt     	   = Request.Form("guardCnt")
		
		c_Time_1159      = "11:59:59 PM"
		
		
		lineupTimeChk = time() - 1/24
		objRSLU.Open "select *  from tbl_lineups where tbl_lineups.GameDay = CDATE('"&gamedate&"') AND tbl_lineups.OwnerID = "& ownerid & "", objConn,3,3,1
		
		'**********************************************************************************************	
		'*** 1) Initialize counters
		'*** 2) Load counter for each position if value is > 0
		'*** 3) Add all counters to determine if the user hit the submit w/o selecting a value		
		'***********************************************************************************************
		centerCBoxCnt       = 0 
		fowardCBoxCnt       = 0 
		foward2CBoxCnt      = 0 
		guardCBoxCnt        = 0
		guard2CBoxCnt       = 0
		int_TotalCboxes     = 0
		tip_time_error_flag = 0
		
		x_cDeadlinePassed   = Request.Form("cDeadlinePassed")
		x_f1DeadlinePassed  = Request.Form("f1DeadlinePassed")
		x_f2DeadlinePassed  = Request.Form("f2DeadlinePassed")
		x_g1DeadlinePassed  = Request.Form("g1DeadlinePassed")
		x_g2DeadlinePassed  = Request.Form("g2DeadlinePassed")
		
		fowardCBoxCnt       = Request.Form("sForward").count
		guardCBoxCnt        = Request.Form("sGuard").count
		
		'###########################
		' Center Logic
		'###########################
		if Request.Form("sCenter").count > 0 then 
			centerCBoxCnt = Request.Form("sCenter").count
			sCenter = Split(Request.Form("sCenter"), ";")	
	
			wCenterName  = sCenter(0)
			wCenterPID   = sCenter(1)
			wCenterBarps = sCenter(2)
			wCenterTip   = sCenter(3)
			
			if CDATE(wCenterTip) < CDATE(lineupTimeChk) AND CDATE(gamedate) = date() then			
				errorCode  = errorCode & "<strong>" & wCenterName & "'s</strong> tip time (" & wCenterTip & ") has passed " & _
				"and he can't be used if not submitted in a previous line-up. - <strong> " & lineupTimeChk & "</strong><br><br>"
	
				errorCodeLU  = "Tip Time Error"
				tip_time_error_flag = 1
			end if
			
		elseif x_cDeadlinePassed = "True" then
			wCenterPID   = objRSLU.Fields("sCenter").Value
			wCenterBarps = objRSLU.Fields("sCenterBarps").Value
			wCenterTip   = objRSLU.Fields("sCenterTip").Value		    			
		else
			wCenterPID   = 9998
			wCenterBarps = 0
			wCenterTip   = c_Time_1159
		end if 

		'###########################
		' Forward Logic
		'###########################
		fowardCBoxCnt = 0
		wForwardPID   = 0
		wForward2PID  = 0
		wCheckForward1= false
		wCheckForward2= false
			
		if x_f1DeadlinePassed = "True" then fowardCBoxCnt = fowardCBoxCnt + 1 end if
		if x_f2DeadlinePassed = "True" then fowardCBoxCnt = fowardCBoxCnt + 1 end if
		fowardCBoxCnt = fowardCBoxCnt + Request.Form("sForward").count
		
		if fowardCBoxCnt > 2 then
			errorCode            = errorCode & "Invalid Line-up - You selected <strong>"& fowardCBoxCnt &"</strong> Forwards.<br>" 
			errorCodeLU          = "Invalid Line-up"
		else
			if x_f1DeadlinePassed= "True" then
				wForwardPID        = objRSLU.Fields("sForward").Value
				wForwardBarps      = objRSLU.Fields("sForwardBarps").Value
				wForwardTip        = objRSLU.Fields("sForwardTip").Value
			end if
			
			if x_f2DeadlinePassed= "True" then
				wForward2PID       = objRSLU.Fields("sForward2").Value
				wForward2Barps     = objRSLU.Fields("sForward2Barps").Value
				wForward2Tip       = objRSLU.Fields("sForwardTip2").Value  
			end if
			
		    lnCount = Request.Form("sForward").count
		    if lnCount > 0 then 			
			    sForward = Split(Request.Form("sForward"),";")	
				
					if lnCount        = 2 then 
						wForwardName    = sForward(0)
						wForwardPID     = sForward(1)
						wForwardBarps   = sForward(2)
						wForwardTip     = sForward(3)
						wForward2Name   = sForward(4)
						wForward2PID    = sForward(5)
						wForward2Barps  = sForward(6)
						wForward2Tip    = sForward(7)					
						wCheckForward1  = true
						wCheckForward2  = true     
					elseif wForwardPID= 0 then
						wForwardName    = sForward(0)
						wForwardPID     = sForward(1)
						wForwardBarps   = sForward(2)
						wForwardTip     = sForward(3)
						wCheckForward1  = true
				else
					wForward2Name     = sForward(0)
					wForward2PID      = sForward(1)
					wForward2Barps    = sForward(2)
					wForward2Tip      = sForward(3)            
					wCheckForward2    = true
				end if
				
				if wCheckForward1 AND CDATE(wForwardTip) < CDATE(lineupTimeChk) AND CDATE(gamedate) = date() then			
					errorCode  = errorCode & "<strong>" & wForwardName & "'s</strong> tip time (" & wForwardTip & ") has passed " & _
					"and he can't be used if not submitted in a previous line-up. - <strong> " & lineupTimeChk & "</strong><br><br>" 
					
					errorCodeLU  = "Tip Time Error"
					tip_time_error_flag = 1
				end if
					
				if wCheckForward2 AND CDATE(wForward2Tip) < CDATE(lineupTimeChk) AND CDATE(gamedate) = date() then			
					errorCode  = errorCode & "<strong>" & wForward2Name & "'s</strong> tip time (" & wForward2Tip & ") has passed " & _
					"and he can't be used if not submitted in a previous line-up. - <strong> " & lineupTimeChk & "</strong><br><br>" 
					
					errorCodeLU  = "Tip Time Error"
					tip_time_error_flag = 1
				end if
			end if
		end if 
		
		
		'Populate Forward 1 with the dummy values if None Selected
		if wForwardPID = 0 then   
			wForwardPID  = 9996
			wForwardBarps= 0
			wForwardTip  = c_Time_1159  
		end if

		'Populate Forward 2 with the dummy values if None Selected
		if wForward2PID 	= 0 then
			wForward2PID   	= 9997
			wForward2Barps 	= 0
			wForward2Tip   	= c_Time_1159   
		end if 
		
		'###########################
		' Guard Logic
		'###########################
		guard2CBoxCnt= 0
		wGuardPID    = 0
		wGuard2PID   = 0
		wCheckGuard1 = false
		wCheckGuard2 = false
			
		if x_g1DeadlinePassed = "True" then guard2CBoxCnt = guard2CBoxCnt + 1 end if
		if x_g2DeadlinePassed = "True" then guard2CBoxCnt = guard2CBoxCnt + 1 end if
		guard2CBoxCnt = guard2CBoxCnt + Request.Form("sGuard").count
		
		if guard2CBoxCnt > 2 then
		    'errorCode      = "Invalid Line-up - Check Box count of <strong>"& fowardCBoxCnt &"</strong> not allowed"
			errorCode      = errorCode & "Invalid Line-up - You selected <strong>"& guard2CBoxCnt &"</strong> Guards." 
			errorCodeLU    = "Invalid Line-up"
		else
		    if x_g1DeadlinePassed = "True" then
			    wGuardPID   = objRSLU.Fields("sGuard").Value
			    wGuardBarps = objRSLU.Fields("sGuardBarps").Value
			    wGuardTip   = objRSLU.Fields("sGuardTip").Value
			end if
			
			if x_g2DeadlinePassed = "True" then
				wGuard2PID             = objRSLU.Fields("sGuard2").Value
			    wGuard2Barps         = objRSLU.Fields("sGuard2Barps").Value
			    wGuard2Tip           = objRSLU.Fields("sGuardTip2").Value
					

			end if
			
		    lnCount = Request.Form("sGuard").count
		    if lnCount > 0 then 			
			    sGuard = Split(Request.Form("sGuard"),";")
					if lnCount      = 2 then 
						wGuardName    = sGuard(0)
						wGuardPID     = sGuard(1)
						wGuardBarps   = sGuard(2)
						wGuardTip     = sGuard(3)
						wGuard2Name   = sGuard(4)
						wGuard2PID    = sGuard(5)
						wGuard2Barps  = sGuard(6)
						wGuard2Tip    = sGuard(7)
						wCheckGuard1  = true
						wCheckGuard2  = true
					elseif wGuardPID = 0 then
						wGuardName    = sGuard(0)
						wGuardPID     = sGuard(1)
						wGuardBarps   = sGuard(2)
						wGuardTip     = sGuard(3)
						wCheckGuard1  = true
					else
						wGuard2Name   = sGuard(0)
						wGuard2PID    = sGuard(1)
						wGuard2Barps  = sGuard(2)
						wGuard2Tip    = sGuard(3)            
						wCheckGuard2  = true
					end if
					
				if wCheckGuard1 AND CDATE(wGuardTip) < CDATE(lineupTimeChk) AND CDATE(gamedate) = date() then			
					errorCode  = errorCode & "<strong>" & wGuardName & "'s</strong> tip time (" & wGuardTip & ") has passed " & _
					"and he can't be used if not submitted in a previous line-up. - <strong> " & lineupTimeChk & "</strong><br><br>" 
		
					errorCodeLU  = "Tip Time Error"
					tip_time_error_flag = 1
				end if
					
				if wCheckGuard2 AND CDATE(wGuard2Tip) < CDATE(lineupTimeChk) AND CDATE(gamedate) = date() then			
					errorCode  = errorCode & "<strong>" & wGuard2Name & "'s</strong> tip time (" & wGuard2Tip & ") has passed " & _
					"and he can't be used if not submitted in a previous line-up. - <strong> " & lineupTimeChk & "</strong>"
					
					errorCodeLU  = "Tip Time Error"
					tip_time_error_flag = 1
				end if				
				
			end if
		end if 
		
		
		'Populate Forward 1 with the dummy values if None Selected
		if wGuardPID = 0 then   
			wGuardPID  = 9994
			wGuardBarps= 0
			wGuardTip  = c_Time_1159  
		end if

		'Populate Forward 2 with the dummy values if None Selected
		if wGuard2PID = 0 then
			wGuard2PID  = 9995
			wGuard2Barps= 0
			wGuard2Tip  = c_Time_1159   
		end if  
		

		int_TotalCboxes = CInt(centerCBoxCnt) + CInt(fowardCBoxCnt) + CInt(guardCBoxCnt) + CInt(foward2CBoxCnt) + CInt(guard2CBoxCnt)
		
		'UNCOMMENT WHEN READY
		
		'x_cDeadlinePassed   = Request.Form("cDeadlinePassed")
		'x_f1DeadlinePassed  = Request.Form("f1DeadlinePassed")
		'x_f2DeadlinePassed  = Request.Form("f2DeadlinePassed")
		'x_g1DeadlinePassed  = Request.Form("g1DeadlinePassed")
		'x_g2DeadlinePassed  = Request.Form("g2DeadlinePassed")
		
		
		if int_TotalCboxes <= 0 and x_cDeadlinePassed = "False" and x_f1DeadlinePassed = "False" and x_f2DeadlinePassed = "False" and x_g1DeadlinePassed = "False" and  x_g2DeadlinePassed = "False"  then
			'errorCode   = "Invalid Line-up - You have selected <strong>"& int_TotalCboxes&"</strong> players to Start in your Line-up."
			errorCode   = "Invalid Line-up - You hit submit without selecting any players. <strong> <br>COME ON MAN!!!</strong>"
			errorCodeLU = "Invalid Line-up" 
		end if
		
		'********************************************************************************************************************
		'***  Check to see if someone is trying to take a player out of today's lineup whose game time has passed.
		'***  Set Error Code when this situation exists 
		'CDATE(wGuardTip) < CDATE(lineupTimeChk
		'********************************************************************************************************************		
		if errorCode          = "" then
		   if objRSLU.RecordCount > 0 AND CDATE(gamedate) = date() then
		      dbTip = objRSLU.Fields("sCenterTip").Value
			  'Response.Write "center Tip  = "&dbTip&".<br>"
		      if CDATE(dbTip) < CDATE(lineupTimeChk) then
		         dbPID = objRSLU.Fields("sCenter").Value				 
		         if dbPID <> wCenterPID then
				    objRSWork.Open "select firstName&' '&lastName as fullname from tblplayers where pid = "&dbPid, objConn
			        dbName = objRSWork.Fields("fullname").Value
				    objRSWork.Close
					
			        errorCode  = errorCode & "<strong>" & dbName & "'s</strong> tip time <strong>(" & dbTip & ")</strong> has passed " & _
		            "and he cannot be removed from your lineup." & _
					"<br>Current time is - <strong> " & lineupTimeChk & "</strong><br><br>"
	
		            errorCodeLU  = "Tip Time Error"
		            tip_time_error_flag = 1
			     end if
		      end if
			  
			  dbTip = objRSLU.Fields("sForwardTip").Value
		      if CDATE(dbTip) < CDATE(lineupTimeChk) then
		         dbPID = objRSLU.Fields("sForward").Value				 
		         if dbPID <> wForwardPID then
				    objRSWork.Open "select firstName&' '&lastName as fullname from tblplayers where pid = "&dbPid, objConn
			        dbName = objRSWork.Fields("fullname").Value
				    objRSWork.Close
					
			        errorCode  = errorCode & "<strong>" & dbName & "'s</strong> tip time <strong>(" & dbTip & ")</strong> has passed " & _
		            "and he cannot be removed from your lineup." & _
					"<br>Current time is - <strong> " & lineupTimeChk & "</strong><br><br>"
	
		            errorCodeLU  = "Tip Time Error"
		            tip_time_error_flag = 1
			     end if
		      end if
			  
			  dbTip = objRSLU.Fields("sForwardTip2").Value
		      if CDATE(dbTip) < CDATE(lineupTimeChk) then
		         dbPID = objRSLU.Fields("sForward2").Value				 
		         if dbPID <> wForward2PID then
				    objRSWork.Open "select firstName&' '&lastName as fullname from tblplayers where pid = "&dbPid, objConn
			        dbName = objRSWork.Fields("fullname").Value
				    objRSWork.Close
					
			        errorCode  = errorCode & "<strong>" & dbName & "'s</strong> tip time <strong>(" & dbTip & ")</strong> has passed " & _
		            "and he cannot be removed from your lineup." & _
					"<br>Current time is - <strong> " & lineupTimeChk & "</strong><br><br>"
	
		            errorCodeLU  = "Tip Time Error"
		            tip_time_error_flag = 1
			     end if
		      end if
			  
			  ' , wGuard2PID
			  dbTip = objRSLU.Fields("sGuardTip").Value
		      if CDATE(dbTip) < CDATE(lineupTimeChk) then
		         dbPID = objRSLU.Fields("sGuard").Value				 
		         if dbPID <> wGuardPID then
				    objRSWork.Open "select firstName&' '&lastName as fullname from tblplayers where pid = "&dbPid, objConn
			        dbName = objRSWork.Fields("fullname").Value
				    objRSWork.Close
					
			        errorCode  = errorCode & "<strong>" & dbName & "'s</strong> tip time <strong>(" & dbTip & ")</strong> has passed " & _
		            "and he cannot be removed from your lineup." & _
					"<br>Current time is - <strong> " & lineupTimeChk & "</strong><br><br>"
	
		            errorCodeLU  = "Tip Time Error"
		            tip_time_error_flag = 1
			     end if
		      end if
			  
			  dbTip = objRSLU.Fields("sGuardTip2").Value
		      if CDATE(dbTip) < CDATE(lineupTimeChk) then
		         dbPID = objRSLU.Fields("sGuard2").Value				 
		         if dbPID <> wGuard2PID then
				    objRSWork.Open "select firstName&' '&lastName as fullname from tblplayers where pid = "&dbPid, objConn
			        dbName = objRSWork.Fields("fullname").Value
				    objRSWork.Close
					
			        errorCode  = errorCode & "<strong>" & dbName & "'s</strong> tip time <strong>(" & dbTip & ")</strong> has passed " & _
		            "and he cannot be removed from your lineup." & _
					"<br>Current time is - <strong> " & lineupTimeChk & "</strong><br><br>"
	
		            errorCodeLU  = "Tip Time Error"
		            tip_time_error_flag = 1
			     end if
		      end if
			  
		   end if
		end if
						
		
		'*******************************************************************************
		'***  Check for Duplicate Players being Selected for multiple positions
		'***  Set Error Code when this situation exists 
		'*******************************************************************************		
		if errorCode          = "" then 
			if (wCenterPID      = wForwardPID) or (wCenterPID  = wForward2PID) then
				'Response.Write "CENTER ERROR <br>"
				errorCode         = "Duplicate Player [<strong>"& wCenterName &"</strong>]!"
				errorCodeLU       = "Invalid Line-up"
			elseif (wForwardPID= wGuardPID) or (wForwardPID= wGuard2PID) then
				'Response.Write "FORWARD ERROR <br>"
				errorCode         = "Duplicate Player [<strong>"& wForwardName &"</strong>]!"
				errorCodeLU       = "Invalid Line-up"    
			elseif (wForward2PID = wGuardPID) or (wForward2PID  = wGuard2PID) then
				'Response.Write "FORWARD ERROR <br>"
				errorCode         = "Duplicate Player [<strong>"& wForward2Name &"</strong>]!"
				errorCodeLU       = "Invalid Line-up"
			end if
		end if
	
		'*******************************************************************************
		'***  Verify the players selected play on the selected date.  
		'*******************************************************************************
		if errorCode  = "" then 
		   FuncCall = Verify_Player_Schedule(wCenterPID, wCenterName)
		   FuncCall = Verify_Player_Schedule(wForwardPID, wForwardName)
		   FuncCall = Verify_Player_Schedule(wForward2PID, wForward2Name)
		   FuncCall = Verify_Player_Schedule(wGuardPID, wGuardName)
		   FuncCall = Verify_Player_Schedule(wGuard2PID, wGuard2Name)

			 if errorCode <> "" then
			 errorCode = errorCode & "<br> This happens when you <strong>DON'T</strong> click the <strong>'Retrieve Lineup'</strong> button upon selecting a new game date."
			 end if		   
		end if
				
		'##########################################################
		' Set to 0 if the player does not have a Barp average
		'##########################################################
		if wCenterBarps < 0 then
			wCenterBarps  = 0
		end if

		if wForwardBarps < 0 then
			wForwardBarps = 0
		end if

		if wForward2Barps < 0 then
			wForward2Barps= 0
		end if

		if wGuardBarps < 0 then
			wGuardBarps   = 0
		end if

		if wGuard2Barps < 0 then
			wGuard2Barps  = 0
		end if  
		
			objRSLU.Close	

		if errorCode  = "" then
			'************************************************************************************
			'****Check now to see if line-up for game date exists
			'****IF fnd process the players
			'****  1) load player from the Record Set
			'****  2) verify if player Tip Time has passed, if so set error code	
			'****  3) Update Player Position
			'****ELSE
			'****  1) Insert Line-up (OwnerID and GameDate)
			'****  2) verify if player Tip Time has passed, if so set error code	
			'****  3) Update Player Position
			'************************************************************************************							

			strSQL = "DELETE FROM tbl_Lineups WHERE tbl_Lineups.GameDay =  #" & gamedate & "#  AND tbl_Lineups.OwnerID = "& ownerid & ";"
			objConn.Execute strSQL

			strSQL ="insert into tbl_lineups(OwnerID,GameDay,sCenter,sCenterBarps,sForward,sForwardBarps,sForward2,sForward2Barps,sGuard,sGuardBarps,sGuard2,sGuard2Barps,sCenterTip,sForwardTip,sForwardTip2,sGuardTip,sGuardTip2) values ('" &_
			ownerid & "', '" &  gamedate & "', '" & wCenterPID & "', '" & wCenterBarps & "', '" & wForwardPID & "', '" & wForwardBarps & "', '" & wForward2PID & "', '" &_
			wForward2Barps & "', '" & wGuardPID & "', '" & wGuardBarps & "', '" & wGuard2PID & "', '" & wGuard2Barps & "', '" &	wCenterTip & "', '" & wForwardTip & "', '" & wForward2Tip & "', '" & wGuardTip & "', '" & wGuard2Tip & "')"
			objConn.Execute strSQL		
			'Response.Write "LINEUP = "&strSQL&".<br>"			
			
			strSQL ="insert into tbl_lineups_history(OwnerID,GameDay,sCenter,sCenterBarps,sForward,sForwardBarps,sForward2,sForward2Barps,sGuard,sGuardBarps,sGuard2,sGuard2Barps,sCenterTip,sForwardTip,sForwardTip2,sGuardTip,sGuardTip2) values ('" &_
			ownerid & "', '" &  gamedate & "', '" & wCenterPID & "', '" & wCenterBarps & "', '" & wForwardPID & "', '" & wForwardBarps & "', '" & wForward2PID & "', '" &_
			wForward2Barps & "', '" & wGuardPID & "', '" & wGuardBarps & "', '" & wGuard2PID & "', '" & wGuard2Barps & "', '" &	wCenterTip & "', '" & wForwardTip & "', '" & wForward2Tip & "', '" & wGuardTip & "', '" & wGuard2Tip & "')"
			objConn.Execute strSQL
			
		end if

		case ""
		ownerid = session("ownerid")	

		if ownerid = "" then
			GetAnyParameter "var_ownerid", ownerid
		end if
		
	case "Return"
		ownerid = session("ownerid")
		
	case "Opponents"
		ownerid = session("ownerid")		
		sURL = "opponents.asp"
		AddLinkParameter "var_ownerid", ownerid, sURL
		Response.Redirect sURL
	
	case "Retrieve Lineup"	
		ownerid = session("ownerid")	

		if ownerid = "" then
			GetAnyParameter "var_ownerid", ownerid
		end if

	end select
	
	'*******************************************************************************************
	'*** NEW CODE FOR STAGGERED LINUPS
	'*******************************************************************************************
	if sAction = "" or sAction = "Retrieve Lineup" then
	
		cDeadlinePassed  = false
		f1DeadlinePassed = false
		f2DeadlinePassed = false
		g1DeadlinePassed = false
		g2DeadlinePassed = false
		selectedCenter   = 0
		selectedForward1 = 0
		selectedForward2 = 0
		selectedGuard1   = 0
		selectedGuard2   = 0
		currentDate      = 0 
		
		if sAction = "Retrieve Lineup" then		
			currentDate = Request.Form("GameDays")
			if currentDate = "" then
				currentDate  = Request.querystring("currentDate")
			end if	
		else
			objRSgames.Open 			"qryGameDeadLines", objConn
			currentDate = CDate(objRSgames.Fields("gameDay"))
			objRSgames.Close
		end if
		
		'*******************************************************************************************
	  '* Subtract 1 hour from the system time.  The system time is EST.  All dates in the
		'* database are stored in CST.
	    '*******************************************************************************************
		lineupTimeChk = time() - 1/24			
  	objRSLU.Open "select *  from tbl_lineups where tbl_lineups.GameDay = CDATE('"&currentDate&"') AND tbl_lineups.OwnerID = "& ownerid & "", objConn,3,3,1
	
	    startingCenterPID = 0
		startingForwardPID = 0
		startingForward2PID = 0
		startingGuardPID = 0
		startingGuard2PID = 0
		
		if objRSLU.RecordCount > 0 then
			lineupFnd = objRSLU.RecordCount
										
			lineupDate          = objRSLU.Fields("gameDay").Value
			startingCenterPID   = objRSLU.Fields("sCenter").Value
			startingCenterTip   = objRSLU.Fields("sCenterTip").Value
			startingForwardPID  = objRSLU.Fields("sForward").Value
			startingForwardTip  = objRSLU.Fields("sForwardTip").Value
			startingForward2PID = objRSLU.Fields("sForward2").Value
			startingForward2Tip = objRSLU.Fields("sForwardTip2").Value			
			startingGuardPID    = objRSLU.Fields("sGuard").Value
			startingGuardTip    = objRSLU.Fields("sGuardTip").Value
			startingGuard2PID   = objRSLU.Fields("sGuard2").Value
			startingGuard2Tip   = objRSLU.Fields("sGuardTip2").Value			
		
			if IsNull(startingCenterTip) then
				cDeadlinePassed = false
			elseif CDATE(startingCenterTip) < CDATE(lineupTimeChk) AND CDATE(currentDate) = date() then 
				cDeadlinePassed = true
				selectedCenter  = objRSLU.Fields("sCenter").Value
			end if
			
			if IsNull(startingForwardTip) then
				f1DeadlinePassed= false
			elseif CDATE(startingForwardTip) < CDATE(lineupTimeChk) AND CDATE(currentDate) = date() then 
				f1DeadlinePassed= true
				selectedForward1= objRSLU.Fields("sForward").Value
			end if
						
			if IsNull(startingForward2Tip) then
				f2DeadlinePassed= false
			elseif CDATE(startingForward2Tip) < CDATE(lineupTimeChk) AND CDATE(currentDate) = date() then 
				f2DeadlinePassed= true
				selectedForward2= objRSLU.Fields("sForward2").Value
			end if
					
			if IsNull(startingGuardTip) then
				g1DeadlinePassed= false   
			elseif CDATE(startingGuardTip) < CDATE(lineupTimeChk) AND CDATE(currentDate) = date() then 
				g1DeadlinePassed= true
				selectedGuard1  = objRSLU.Fields("sGuard").Value		
			end if

			if IsNull(startingGuard2Tip) then
				g2DeadlinePassed= false
			elseif CDATE(startingGuard2Tip) < CDATE(lineupTimeChk) AND CDATE(currentDate) = date() then 
				g2DeadlinePassed= true
				selectedGuard2  = objRSLU.Fields("sGuard2").Value
			end if

		end if
		
		'*******************************************************************************************
		'*** END NEW CODE FOR STAGGERED LINUPS
		'*******************************************************************************************	
	end if
	sURL = ".asp"
	
	objRSteamRec.Open "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & ownerid & "))", objConn
	w_tm_rank = objRSteams.Fields("rank").value
	
	currentDate = date()
	objRSgames.Open  "select * from qryGameDeadLines", objConn,1,1
	
	if sAction  = "Retrieve Lineup" then  
		selectDate= Request.Form("GameDays")
		if selectDate = "" then 
			selectDate = Request.querystring("currentDate")
		end if	
	else
	    if objRSgames.RecordCount > 0 then
		    selectDate= CDate(objRSgames.Fields("gameDay"))
		else
		    selectDate = currentDate
		end if
	end if
	

	objRSteams.Open    "SELECT TeamName FROM qryTeams WHERE qryTeams.OwnerID        = "&ownerid&" ", objConn
	objRSgameDate.Open "SELECT * FROM tblGameDeadLines where gameday = cdate('"&currentDate&"') ", objConn,1,1
							  				  
	if cDate(selectDate) = date() then
		objRSGameStarted.Open "SELECT * FROM tblPlayers t WHERE ownerid = "&ownerid&" and exists " &_
													"(select 1 from NBAINDTMSked s where s.NBATeam = t.NBATEAMID and s.gameDay = #"&selectDate&"# " &_
													"and s.gametime < #"&lineupTimeChk&"#)",objConn,1,1
						   
		gameStartedCnt = objRSGameStarted.RecordCount
		objRSGameStarted.Close
	else
		gameStartedCnt = 0
	end if
	

  Set objRSAll	= Server.CreateObject("ADODB.RecordSet")
  objRSAll.Open  	"SELECT * FROM qryAllGames where AwayTeamInd = " & ownerid & " " &_
                            "OR HomeTeamInd  = " & ownerid & " ", objConn
  oppcnt = 0
	
  '*****************************************************************************
  ' Verify_Player_Schedule()
  ' Verify that the player plays on selected date.  An error could occur if the 
  ' user selected a date and but did not select retrieve lineup
  ' prior to submitting the lineup.
  '****************************************************************************
  Function Verify_Player_Schedule (p_PlayerID, p_PlayerName)
	 if p_PlayerID < 9000 then
	    objRSWork.Open 		"SELECT * " & _
												"FROM tblPlayers t, nbaindtmsked n " & _
												"WHERE  PID = "& p_PlayerID & "  " & _
												"AND t.NBATeamID = n.NBATeam " & _
												"AND n.GameDay = #"& gamedate & "# ", objConn,3,3,1

	    if objRSWork.RecordCount = 0 then
		   wPlayerNm = Replace (p_PlayerName, ",", "")
		   errorCode = errorCode & "<strong>"& wPlayerNm &"</strong> does not play on "&gamedate&".<br>"
			 errorCodeLU = "Invalid Line-up"		
	    end if
		
		objRSWork.close
	 end if

   End Function
  
%>

<%		
	'**********************************************************************************************'
	'**** EDIT PLAYER SECTION
	'**********************************************************************************************'

	if Request.Form("action") <> "Save Form Data" Then
		'Nothing
	else
		
		password              = Request.Form("txtPassword")
		confirmation          = Request.Form("confirmation")
		userName              = Request.Form("txtUserName")
		ownerName             = Request.Form("txtownerName")
		teamName              = Request.Form("txtTeamName")
		hPhone                = Request.Form("txtHPhone")
		wPhone                = Request.Form("txtWPhone")
		cPhone                = Request.Form("txtCPhone")
		hEmail                = Request.Form("txtEmailH")
		textMessages          = Request.Form("textMessages")
		
		if Request.Form("receiveFreeAgentAlerts").count > 0 then 
			receiveFreeAgentAlerts = true
		else
			receiveFreeAgentAlerts = false
		end if
		
		if Request.Form("receiveTradeAlerts").count > 0 then 
			receiveTradeAlerts = true
		else
			receiveTradeAlerts = false
		end if

		if Request.Form("receiveWaiverAlerts").count > 0 then 
			receiveWaiverAlerts = true
		else
			receiveWaiverAlerts = false
		end if

		if Request.Form("receiveStaggerAlerts").count > 0 then 
			receiveStaggerAlerts = true
		else
			receiveStaggerAlerts = false
		end if

		if Request.Form("receiveBoxScoreAlerts").count > 0 then 
			receiveBoxScoreAlerts = true
		else
			receiveBoxScoreAlerts = false
		end if

		if Request.Form("receiveOTBAlerts").count > 0 then 
			receiveOTBAlerts = true
		else
			receiveOTBAlerts = false
		end if
		
		if Request.Form("receiveRentalAlerts").count > 0 then 
			receiveRentalAlerts   = true
		else
			receiveRentalAlerts   = false
		end if	

		if Request.Form("acceptTradeOffers").count > 0 then 
			acceptTradeOffers   = true
		else
			acceptTradeOffers   = false
		end if			

		if Request.Form("receiveEmails").count > 0 then 
			receiveEmails   = true
		else
			receiveEmails   = false
		end if	

		if Request.Form("receiveTexts").count > 0 then 
			receiveTexts   = true
		else
			receiveTexts   = false
		end if			

		if Request.Form("receiveEmailLeagueAlerts").count > 0 then 
			receiveEmailLeagueAlerts   = true
		else
			receiveEmailLeagueAlerts   = false
		end if			
				
		shortname             = Request.Form("txtShortName")
		teamlogo              = Request.Form("txtteamlogo")

 		if   confirmation = password then       
  		strSQL ="update tblOwners Set textMessages =  '" & textMessages  & "',Password =  '" & password  & "',acceptTradeOffers =  " & acceptTradeOffers  & ",receiveEmailLeagueAlerts =  " & receiveEmailLeagueAlerts  & ",receiveFreeAgentAlerts =  " & receiveFreeAgentAlerts  & ",receiveTradeAlerts =  " & receiveTradeAlerts  & ",receiveWaiverAlerts =  " & receiveWaiverAlerts  & ",receiveStaggerAlerts =  " & receiveStaggerAlerts  & ",receiveBoxScoreAlerts =  " & receiveBoxScoreAlerts  & ",receiveRentalAlerts =  " & receiveRentalAlerts  & ",receiveOTBAlerts =  " & receiveOTBAlerts  & ",receiveTexts =  " & receiveTexts  & ",receiveEmails =  " & receiveEmails  & ",passconf =  '" & confirmation  & "',userID = '" & username & "',OwnerName = '" & ownername & "',TeamName= '" & teamname & "',CellPhone = '" & cphone & "',HomeEmail = '" & hemail &"', ShortName = '" & shortname & "' , teamlogo = '" & TeamLogo & "' where tblOwners.OwnerID = " & ownerid & ";"
			objConn.Execute strSQL		
			strSQL = "update tblOwners set UpdateTime = now() - 1/24 where ownerid = "& ownerid & ";"
			objConn.Execute strSQL			
			sURL = "dashboard.asp"
			AddLinkParameter "var_ownerid", ownerid, sURL
	    Response.Redirect sURL
  		else
  			 Response.Write "<H4>Password and Password Confirmation Don't Match: RE-ENTER PASSWORD! </H4> "
		end if
 end if
 		
%>
<%	
'**********************************************************************************************'
'**** OTB SECTION
'**********************************************************************************************' 
	
	GetAnyParameter "Action", sAction	

	if sAction = "Update" then
		'Do Nothing
	else	
		TEAM_Split        = Split(Request.Form("Action"), ";")
		sAction           = TEAM_Split(0) 'SAction
		tradepartner      = TEAM_Split(1) 'Team ID
		sTeam             = TEAM_Split(1) 'Team			
	end if	
	
	'Response.Write "IN OTB   " &sAction& "<br> "
	
	select case sAction

	case "Continue"
		sURL = "tradeanalyzer.asp"
		conAction = "Continue"
		AddLinkParameter "Action", conAction, sURL
		AddLinkParameter "cmbTeam", tradepartner, sURL
		Response.Redirect sURL
		
	end select

	if sAction = "Update" then
		'##############################
		'# Update values on tblowners
		'##############################
		blockComments = Request.Form("ontheblockscomments")
		nCenters  = Request.Form("chkCenters").Count
		nForwards = Request.Form("chkForwards").Count
		nGuards   = Request.Form("chkGuards").Count
		nForCen   = Request.Form("chkFC").Count
		nGuardFor = Request.Form("chkGF").Count
		nAllPos   = Request.Form("chkAllPlayers").Count

		strSQL = 	"UPDATE tblowners " &_
							"SET ontheblockcomments = '"&blockComments&"', ontheblockneedscen = "&nCenters&", ontheblockall = "&nAllPos	&",  " &_
							"ontheblockneedsfor = "&nForwards&", ontheblockneedsgua = "&nGuards&",  " &_
							"ontheblockneedsfc =  "&nForCen&", ontheblockneedsgf = "&nGuardFor&"  " &_
							"WHERE ownerid = "&ownerid&" "
									
		objConn.Execute strSQL
			 
	 	'##################################
		'# Reset all players for this owner
		'##################################
		strSQL = "UPDATE tblPlayers SET ontheblock = 0 " &_
							"WHERE ownerid = "&ownerid&" "
		objConn.Execute strSQL

		'################################################################
		'# Loop thru all checked players and set the ontheblock indicator
		'################################################################
		For I = 1 To Request.Form("chkPID").Count
		PID_Split = Split(Request.Form("chkPID")(I), ";")
		PID_Split(0)
		PID_Split(1)
		iRecordToUpdate = PID_Split(0)

			strSQL = "UPDATE tblPlayers SET ontheblock = 1 " &_
							 "WHERE pid = "&iRecordToUpdate&" "
			objConn.Execute strSQL
		NEXT

			email_test = Request.Form("email_league")

			if email_test = "Yes" then

				'*****************************************************
				'EMAIL TO THE LEAGUE 
				'*****************************************************
				Set objRSteams = Server.CreateObject("ADODB.RecordSet")
				Set objEmail   = Server.CreateObject("ADODB.RecordSet")
				Set objrsNames = Server.CreateObject("ADODB.RecordSet")
			
				objRSteams.Open "SELECT * FROM tblOwners where ownerid = "& ownerid & "", objConn
				ontheblock_owner_name= objRSteams.Fields("ShortName").Value
				objRSteams.Close
				
				objrsNames.Open "SELECT * FROM tblPlayers Where ownerID = "&ownerid&" and onTheBlock = true " , objConn
				otbPlayers = null
				While Not objrsNames.EOF
					otbPlayers = otbPlayers & objrsNames.Fields("firstName").Value & " " & objrsNames.Fields("lastName").Value & "<br>"  
					objrsNames.MoveNext
				Wend
				objrsNames.close
					
								wEmailOwnerID = null					
				wAlert        = "receiveOTBAlerts"
				email_subject = ontheblock_owner_name&" Has Updated the Block!"
				email_message = otbPlayers & "<br>" & blockComments
%>		
		<!--#include virtual="Common/email_league.inc"-->
<%		
			end if
			sURL = "dashboard.asp"
			AddLinkParameter "var_ownerid", ownerid, sURL
	    Response.Redirect sURL
 	end if	

%>	
<!--#include virtual="Common/noTrades.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
red {
	color:#9a1400;
	font-weight:bold;
}

black {
	color:black;
}
h4 {
color:#01579B;
}
.alert-warning {
    color: #01579B;
    background-color:#ddd;
    border-color:#01579B;
}
.jumbotron {
    background-color: #9a1400;
    color: white;
    margin-bottom: 15px;
		font-weight:bold;
}
span.blackIcon{
	color:black;
	font-weight: bold;
	text-transform: none;
}		
greenIcon2{
	color:#468847;
	text-transform: none;
}	
redIcon2{
	color:#9a1400;
	text-transform: none;
}
blackIcon2{
	color:black;
	text-transform: none;
}	

.alert-danger {
    color: #ffffff;
    background-color: #9a1400;
    border-color: #111;
}
.nav-tabs {
    border-bottom: 2px solid black;
}
p {
border: 1px solid red;
padding: 5px;
background-color:beige;
}
red {
color: red;
font-weight:bold;
}
.item {
background: #333;
text-align: center;
height: 120px !important;
}
.panel-primary {
background-color: #194719 !important;
}
td{
	text-align: center;
	vertical-align: middle;
}

</style>
</head>
<body>
<!--#include virtual="Common/headerMain.inc"-->
<script>
function toggle(source) {
  checkboxes = document.getElementsByName('chkPid');
  for(var i=0, n=checkboxes.length;i<n;i++) {
    checkboxes[i].checked = source.checked;
  }
}
	$.fn.pageMe = function(opts){
	var $this = this,
	defaults = {
	perPage: 7,
	showPrevNext: false,
	hidePageNumbers: false
	},
	settings = $.extend(defaults, opts);

	var listElement = $this;
	var perPage = settings.perPage;
	var children = listElement.children();
	var pager = $('.pager');

	if (typeof settings.childSelector!="undefined") {
	children = listElement.find(settings.childSelector);
	}

	if (typeof settings.pagerSelector!="undefined") {
	pager = $(settings.pagerSelector);
	}

	var numItems = children.size();
	var numPages = Math.ceil(numItems/perPage);

	pager.data("curr",0);

	if (settings.showPrevNext){
	$('<li><a href="#" class="prev_link">«</a></li>').appendTo(pager);
	}

	var curr = 0;
	while(numPages > curr && (settings.hidePageNumbers==false)){
	$('<li><a href="#" class="page_link">'+(curr+1)+'</a></li>').appendTo(pager);
	curr++;
	}

	if (settings.showPrevNext){
	$('<li><a href="#" class="next_link">»</a></li>').appendTo(pager);
	}

	pager.find('.page_link:first').addClass('active');
	pager.find('.prev_link').hide();
	if (numPages<=1) {
	pager.find('.next_link').hide();
	}
	pager.children().eq(1).addClass("active");

	children.hide();
	children.slice(0, perPage).show();

	pager.find('li .page_link').click(function(){
	var clickedPage = $(this).html().valueOf()-1;
	goTo(clickedPage,perPage);
	return false;
	});
	pager.find('li .prev_link').click(function(){
	previous();
	return false;
	});
	pager.find('li .next_link').click(function(){
	next();
	return false;
	});

	function previous(){
	var goToPage = parseInt(pager.data("curr")) - 1;
	goTo(goToPage);
	}

	function next(){
	goToPage = parseInt(pager.data("curr")) + 1;
	goTo(goToPage);
	}

	function goTo(page){
	var startAt = page * perPage,
	endOn = startAt + perPage;

	children.css('display','none').slice(startAt, endOn).show();

	if (page>=1) {
	pager.find('.prev_link').show();
	}
	else {
	pager.find('.prev_link').hide();
	}

	if (page<(numPages-1)) {
	pager.find('.next_link').show();
	}
	else {
	pager.find('.next_link').hide();
	}

	pager.data("curr",page);
	pager.children().removeClass("active");
	pager.children().eq(page+1).addClass("active");

	}
	};

	$(document).ready(function(){

	$('#myTable').pageMe({pagerSelector:'#myPager',showPrevNext:true,hidePageNumbers:false,perPage:30});

	});
	
$(document).ready(function() {
    $('#example').DataTable( {
        "order": [[ 0, "desc" ]],
			  "lengthMenu": [[25, 50, 75, -1], [25, 50, 75, "All"]]
    } );
} );
</script>
<script src="js/lineup.js"></script>
<%
 	objRSteams.Close
 	Set objRSteams= Nothing
	if sAction = "" or sAction = "Retrieve Lineup" then
	

	   objRSUnavail.Open "SELECT t.*, IIf(b.barps Is Null,0,b.barps) AS barps, tblNBATeams.TeamName,b.usage " &_
	                  "FROM (tblPlayers AS t LEFT JOIN tbl_barps AS b ON (t.lastName = b.last) AND (t.firstName = b.first)) " &_
				  	        "LEFT JOIN tblNBATeams ON t.NBATeamID = tblNBATeams.NBATID " &_
										"WHERE t.ownerid = "&ownerid&" AND " &_
	                  "NOT Exists (select 1 from NBAINDTMSked s where s.NBATeam = t.NBATEAMID and s.gameDay = #"&selectDate&"#) " &_					
										"order by b.barps DESC, t.l5barps DESC, t.lastName",objConn,1,1
										
	   unAvailCnt  = objRSUnavail.RecordCount										
%>
<div class="container">
	<div class="row">
		<div class="col-xs-12">
			<table class="col-xs-8 table table-custom-black table-condensed table-bordered table-responsive">
				<tr style="background-color:white;text-align:center;">									
					<td class="big" style="width: 25%;background-color: yellowgreen;color: white;font-weight: bold;"><span style="color:black;">REG:&nbsp;</span> <%=objRSteamsSession("won")%>-<%=objRSteamsSession("loss")%></td>
					<td class="big" style="width: 25%;background-color: yellowgreen;color: white;font-weight: bold;">CYL:&nbsp;<span style="color:black;">[<%=w_current_cycle%>]</span> <%=cycleWins%>-<%=cycleLoss %></td>
					<td class="big" style="width: 25%;background-color: yellowgreen;color: white;font-weight: bold;"><span style="color:black;">WBAL:&nbsp;</span>$<%= w_WaiverBal%></td>
					<td class="big" style="width: 25%;background-color: yellowgreen;color: white;font-weight: bold;"><span style="color:black;"><i class="fas fa-usd-circle fa-lg"></i>&nbsp;<a href="finreports.asp?ownerid=<%= ownerid %>" style="color: white;text-decoration: underline;">$<%= w_total_spent %></a></td>
					</tr>
			</table>	
		</div>
	</div> 		
</div>
</br>
 <div class="container">
  <ul class="nav nav-tabs">
    <li class="active"><a data-toggle="tab" href="#home"><i class="fas fa-tachometer-alt"></i>&nbsp;Dashboard</a></li>
    <li><a data-toggle="tab" href="#team"><i class="fas fa-user-cog"></i>&nbsp;Edit Team Info</a></li>
    <li><a data-toggle="tab" href="#otb"><i class="fas fa-truck-moving"></i>&nbsp;OTB</a></li>
  </ul>
<!--Beginning of Tabs-->
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<div id="home" class="tab-pane fade in active">					
					<br>	
					<div class="row">
					<!--HEADER INFO & HYPERLINKS-->					
						<div class="col-xs-4">
							<img class="img-responsive pull-left" style="max-width:100%;height:auto;margin:0px auto;display:block;border: #111;border-width: thin;border-style: solid;" src="<%=Logo%>">
						</div>
						<div class="col-xs-8">
							<table class="col-xs-8 table table-custom-black table-condensed table-bordered table-responsive">
								<tr style="background-color:white;">									
									<td class="txn big" style="background-color:white;text-align:left;width:25%"><i class="fas fa-user-plus"></i>&nbsp;<a class="blue" href="transelect.asp">Adds</a></td>
									<td class="txn big" style="background-color:white;text-align:left;width:25%"><i class="far fa-list-ul"></i>&nbsp;<a class="blue" href="allskeds.asp">SKED</a></td>
									<td class="big" style="width:50%;background-color: black;color: dodgerblue;font-weight: bold;text-align:left;">
										<% if  objRStradesreceived.RecordCount > 0 then %>
									<i class="far fa-inbox-in" style="color: yellow;"></i>&nbsp;Trades&nbsp;<a href="tradeoffers.asp?ownerid=<%= ownerid %>" style="color:white;text-decoration:underline;font-size=12px;"><%= objRStradesreceived.RecordCount %></a></button>
										<%else%>
										<span style="color: #616161;"><i class="far fa-inbox-in" ></i>&nbsp;Trades&nbsp;0</span>
										<%end if%>									
									</td>
								</tr>
								<tr>									
									<td class="txn big" colspan="2" style="background-color:white;text-align:left;"><i class="far fa-ticket-alt"></i>&nbsp;<a class="blue" href="viewLineups.asp">Matchups</a></td>
									<td class="big" style="width: 50%;background-color: black;color: yellowgreen;font-weight: bold;text-align:left;">
										<% if objRStradesmade.RecordCount > 0 then %>
										<i class="far fa-inbox-out" style="color: yellow;"></i>&nbsp;Trades&nbsp;<a href="pendingtrades.asp?ownerid=<%= ownerid %>" style="color:white;text-decoration:underline;font-size=12px;"><%= objRStradesmade.RecordCount %></a></button>
										<%else%>
										<span style="color: #616161;"><i class="far fa-inbox-out"></i>&nbsp;Trades&nbsp;0</span>
										<%end if%>									
									</td>
								</tr>
								<tr style="background-color:white">
									<td colspan="2" width="50%" class="txn big text-left"><i class="fal fa-chart-line"></i>&nbsp;<a class="blue" href="teamanalysis.asp">Team Analysis</a></td>
									<td class="big" style="width: 50%;background-color: black;color: turquoise;font-weight: bold;text-align:left;">
										<% if  objRSwaivers.RecordCount > 0 then %>
										<i class="fas fa-user-clock" style="color: yellow;"></i>&nbsp;Waivers&nbsp;<a href="pendingwaivers.asp?ownerid=<%= ownerid %>" style="color:white;text-decoration:underline;font-size=12px;"><%= objRSwaivers.RecordCount %></a></button>
										<%else%>
										<span style="color: #616161;"><i class="fas fa-user-clock"></i>&nbsp;Waivers&nbsp;0</span>
										<%end if%>									
									</td>
								</tr>
								<tr style="background-color:white">
									<td colspan="2" class="txn big text-left"><i class="fal fa-sort-numeric-up"></i>&nbsp;<a class="blue" href="allstandings.asp">Standings</a></span></td>
									<td class="big" style="width: 50%;background-color: black;color: darkorange;font-weight: bold;text-align:left;">
									<% if  objrsForecast.RecordCount > 0 then %>
										<i class="fa fa-balance-scale" style="color: yellow;"></i>&nbsp;Analyzing&nbsp;<a href="pendingAnalyzedTrades.asp?ownerid=<%= ownerid %>" style="color:white;text-decoration:underline;font-size=12px;"><%= objrsForecast.RecordCount %></a></button>
										<%else%>
										<span style="color: #616161;"><i class="fa fa-balance-scale"></i>&nbsp;Analyzing&nbsp;0</span>
										<%end if%>
									</td>										
								</tr>
								<tr style="background-color:white">
									<% if (date() > wPlayoffStart) then %>
									 <td colspan="2" style="width:50%" class="txn big text-left"><i class="fal fa-calculator"></i>&nbsp;<a class="blue" href="results.asp">IGBL Scores</a></td><td class="txn big text-left"><i class="fas fa-users"></i>&nbsp;<a class="blue" href="playerstats.asp">Player Stats</a></td>
									<% else %>
									 <td colspan="2" style="width:50%" class="txn big text-left"><i class="fal fa-calculator"></i>&nbsp;<a class="blue" href="results.asp">IGBL Scores</a></td><td class="txn big text-left"><i class="fas fa-users"></i>&nbsp;<a class="blue" href="playerstats.asp">Player Stats</a></td>
									<% end if %>
								</tr>
								<%if gameStartedCnt > 0 then %>
								<tr style="background-color:white">
									<td colspan="3">
										<table class="table table-responsive table-condensed table-bordered" style="background-color:#ddd;">
											<tr>
												<td style="width:33%;text-align:center"><greenIcon2><i class="fas fa-sync fa-spin"></i> Active<span class="sr-only"></span></greenIcon2></td>
												<td style="width:34%;text-align:center"><redIcon2><i class="fas fa-user-slash"></i>&nbsp;Banned</redIcon2></td>
												<td style="width:33%;text-align:center"><blackIcon2><i class="fas fa-user-lock"></i> Locked</blackIcon2></td>					
											</tr>
										</table>	
									</td>
								</tr>
								<% end if %>

							</table>	
						</div>
					</div>
				<form action="dashboard.asp" method="POST" onSubmit="return functionLineup(this)" name="frmLineups"  language="JavaScript">
					<input type="hidden" name="var_ownerid"  value="<%= ownerid %>" />
					<input type="hidden" name="var_gameTime" value="<%= gameDeadline %>" />
					<input type="hidden" name="var_gameTimeStagger" value="<%= gameStaggerDeadline %>" />
					<input type="hidden" name="selectDate" value="<%= selectDate %>" />
					<input type="hidden" name="centerCnt" value="<%= centerCnt %>" />
					<input type="hidden" name="forwardCnt" value="<%= forwardCnt %>" />
					<input type="hidden" name="guardCnt" value="<%= guardCnt %>" />	
					<input type="hidden" name="cDeadlinePassed"  value="<%= cDeadlinePassed %>" />
					<input type="hidden" name="f1DeadlinePassed" value="<%= f1DeadlinePassed %>" />
					<input type="hidden" name="f2DeadlinePassed" value="<%= f2DeadlinePassed %>" />	
					<input type="hidden" name="g1DeadlinePassed" value="<%= g1DeadlinePassed %>" />
					<input type="hidden" name="g2DeadlinePassed" value="<%= g2DeadlinePassed %>" />
					<input type="hidden" name="selectedCenter" value="<%= selectedCenter %>" />
					<input type="hidden" name="selectedForward1" value="<%= selectedForward1 %>" />
					<input type="hidden" name="selectedForward2" value="<%= selectedForward2 %>" />
					<input type="hidden" name="selectedGuard1" value="<%= selectedGuard1 %>" />
					<input type="hidden" name="selectedGuard2" value="<%= selectedGuard2 %>" />
					</br>
					<div class="row">
						<div class="col-xs-6">
							<select class="form-control " name="gameDays">
								<%if sAction = "Retrieve Lineup" then%>
								<option value="<%=selectDate%>"- selected><%=selectDate%></option>
								<% else %>
								<option value="<%=objRSgames.Fields("gameDay")%>" selected><%=objRSgames.Fields("gameDay")%></option>
								<% end if %>
								<% While not objRSgames.EOF %>
								<option value="<%=objRSgames("gameDay")%>"><%=objRSgames.Fields("gameDay")%> </option>
								<% objRSgames.MoveNext
									Wend 
								%>
							</select>
						</div>
						<div class="col-xs-6">
							<button class="btn btn-default  btn-block" type="submit"  id="RetrieveLineup" value="Retrieve Lineup" name="Action"><span class="glyphicon glyphicon-search"></span>&nbsp;Retrieve Lineup</button>
						</div>
					</div>
					</br>
					<!--TRADELINE DEADLINE BANNER-->
					<% if (date() >= (wTradeDeadLine - 10)) AND (wTradeDeadLine >= date()) then %>
					<div class="row">		
						<div class="col-xs-12">
							<div class="alert alert-danger">
								<strong><i class="fa fa-exclamation-triangle" aria-hidden="true"></i>Trade-Deadline: </strong> <%=wTradeDeadLine%>
							</div>
						</div>
					</div>
					<% end if  				
					objsrAvailPlayers.Open "SELECT * FROM qLineupDaily WHERE (POS = 'CEN' Or POS = 'F-C') and qLineupDaily.OwnerID  = "&ownerid&" and gameday = cdate('"&selectDate&"')",  objConn,1,1
					centerCnt = objsrAvailPlayers.RecordCount
					objsrAvailPlayers.CLOSE
					
					if centerCnt > 0 Then %>		
					<div class="row">
						<div class="col-xs-12">
							<table class="table table-custom-black table-bordered table-responsive table-condensed">
								<tr>
									<th class="text-uppercase text-left;" colspan="2">CENTER</th>
									<th style="text-align:center;">AVG</th>
									<th style="text-align:center;">L/5</th>
									<th style="text-align:center;">USG</th>
									<th style="text-align:center;"><i class="fas fa-basketball-hoop"></i></th>
								</tr>
								<%
					          loopCt = 1
					          while loopCt <= 2 
										if loopCt = 1 then
 				                    objsrAvailPlayers.Open  "SELECT * FROM qLineupDaily WHERE (POS = 'CEN' Or POS = 'F-C') " &_
					                                          "AND qLineupDaily.OwnerID  = "&ownerid&" AND gameday = cdate('"&selectDate&"') " & _
					  	                                      "AND PID = "&startingCenterPID,  objConn,1,1
										else
 				                    objsrAvailPlayers.Open  "SELECT * FROM qLineupDaily WHERE (POS = 'CEN' Or POS = 'F-C') " &_
					                                          "AND qLineupDaily.OwnerID  = "&ownerid&" AND gameday = cdate('"&selectDate&"') " & _
					  	                                      "AND PID <> "&startingCenterPID,  objConn,1,1								 
										end if
								   
										While Not objsrAvailPlayers.EOF
										 if len(objsrAvailPlayers.Fields("GameTime").Value) = 10 then
												wtime = Left(objsrAvailPlayers.Fields("GameTime").Value,4) & Right(objsrAvailPlayers.Fields("GameTime").Value,3)
										 else
												wtime = Left(objsrAvailPlayers.Fields("GameTime").Value,5) & Right(objsrAvailPlayers.Fields("GameTime").Value,3)
										 end if
								%>
									 <%if objsrAvailPlayers.Fields("PID").Value = selectedCenter then%>	
										<tr class="success">
									<%else%>
										<tr style="background-color:white;">
									<%end if%>
									<!--#include virtual="Common/availPlayers.inc"-->		
									<%if cDeadlinePassed = true then %>
									   <%if objsrAvailPlayers.Fields("PID").Value = startingCenterPID then%>	
											<td class="success" style="vertical-align:middle;text-align:center;width:10%"><greenIcon2><i class="fas fa-sync fa-spin"></i><greenIcon2></td>
										<%else%>
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><redIcon2><i class="fas fa-user-slash"></i></redIcon2></td>
										<%end if%>
									<%elseif  ((CDATE(objsrAvailPlayers.Fields("GameTime").Value) <  CDATE(lineupTimeChk)) and (CDATE(selectDate) = date()))then %>
										<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><blackIcon><i class="fas fa-user-lock"></i></blackIcon></td>									
									<%elseif objsrAvailPlayers.Fields("PID").Value = startingCenterPID then %>							
										<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" checked id="C" name="sCenter" value="<%=objsrAvailPlayers.Fields("firstName").Value & " " & objsrAvailPlayers.Fields("lastName").Value & ";" & objsrAvailPlayers.Fields("PID").Value & ";" & objsrAvailPlayers.Fields("barps").Value  & ";" &  objsrAvailPlayers.Fields("GameTime").Value & ";"%>"></td>
									<%else%>
										<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" id="C" name="sCenter" value="<%=objsrAvailPlayers.Fields("firstName").Value & " " & objsrAvailPlayers.Fields("lastName").Value & ";" & objsrAvailPlayers.Fields("PID").Value & ";" & objsrAvailPlayers.Fields("barps").Value  & ";" &  objsrAvailPlayers.Fields("GameTime").Value & ";"%>"></td>
									<%end if%>
								</tr>
									
								<%
									objsrAvailPlayers.MoveNext
									Wend
									objsrAvailPlayers.close
									loopCt = loopCt + 1
									Wend
								%>
							</table>
							<br>
						</div>
					</div>				
					<% else %>
					<div class="row">
						<div class="col-xs-12">
							<div class="jumbotron" style="background-color:#9a1400;color:white;text-align:center;vertical-align:middle">
								<i class="fas fa-frown big"></i>&nbsp;No Centers Available Tonight!
							</div>
						</div>
					</div>
					<% end if 					
					
					objsrAvailPlayers.Open      "SELECT * FROM qLineupDaily WHERE (POS = 'FOR' Or POS = 'F-C' Or POS = 'G-F') and qLineupDaily.OwnerID  = "&ownerid&" and gameday = cdate('"&selectDate&"')",  objConn,1,1
					forwardCnt = objsrAvailPlayers.RecordCount
					objsrAvailPlayers.CLOSE
					
					if forwardCnt > 0 Then %>
					<div class="row">
						<div class="col-xs-12">
							<table class="table table-custom-black table-bordered table-responsive table-condensed">
								<tr>
									<th class="text-uppercase text-left;" colspan="2">FORWARDS</th>
									<th style="text-align:center;">AVG</th>
									<th style="text-align:center;">L/5</th>
									<th style="text-align:center;">USG</th>
									<th style="text-align:center;"><i class="fas fa-basketball-hoop"></i></th>
								</tr>
									<%
	       
					          loopCt = 1
					          while loopCt <= 2 
										if loopCt = 1 then
 				                   objsrAvailPlayers.Open   "SELECT * FROM qLineupDaily WHERE (POS = 'FOR' Or POS = 'F-C' Or POS = 'G-F') " &_
					                                          "AND qLineupDaily.OwnerID  = "&ownerid&" AND gameday = cdate('"&selectDate&"') " & _
					  	                                      "AND PID in ("&startingForwardPID&","&startingForward2PID&")",  objConn,1,1
										else
 				                     objsrAvailPlayers.Open "SELECT * FROM qLineupDaily WHERE (POS = 'FOR' Or POS = 'F-C' Or POS = 'G-F') " &_
					                                          "AND qLineupDaily.OwnerID  = "&ownerid&" AND gameday = cdate('"&selectDate&"') " & _
					  	                                      "AND PID Not in ("&startingForwardPID&","&startingForward2PID&")",  objConn,1,1
										end if
								 
										While Not objsrAvailPlayers.EOF
										 if len(objsrAvailPlayers.Fields("GameTime").Value) = 10 then
												wtime = Left(objsrAvailPlayers.Fields("GameTime").Value,4) & Right(objsrAvailPlayers.Fields("GameTime").Value,3)
										 else
												wtime = Left(objsrAvailPlayers.Fields("GameTime").Value,5) & Right(objsrAvailPlayers.Fields("GameTime").Value,3)
										 end if							   
										 objRSNext5.Open  	"SELECT * FROM qryAllPlayerGameDays WHERE PID = " & objsrAvailPlayers.Fields("PID").Value & " AND gameday >= Date() order by gameday ", objConn,1,1	
									%>
									<%if ((objsrAvailPlayers.Fields("PID").Value = selectedForward1) or (objsrAvailPlayers.Fields("PID").Value = selectedForward2)) then %>
										<tr class="success">
									<%else%>
										<tr style="background-color:white;">
									<%end if%>
										<!--#include virtual	="Common/availPlayers.inc"-->
										<%if f1DeadlinePassed AND f2DeadlinePassed then%>
										  <%if ((objsrAvailPlayers.Fields("PID").Value = selectedForward1) or (objsrAvailPlayers.Fields("PID").Value = selectedForward2)) then %>
												<td class="success" style="vertical-align:middle;text-align:center;width:10%"><greenIcon2><i class="fas fa-sync fa-spin"></i></greenIcon2></td>
											<%else%>
												<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><redIcon2><i class="fas fa-user-slash"></i></redIcon2></td>
											<%end if%>		
										<%elseif ((objsrAvailPlayers.Fields("PID").Value = selectedForward1) or (objsrAvailPlayers.Fields("PID").Value = selectedForward2))then %>
											<td class="success" style="vertical-align:middle;text-align:center;width:10%"><greenIcon2><i class="fas fa-sync fa-spin"></i></greenIcon2></span></td>				
										<%elseif ((CDATE(objsrAvailPlayers.Fields("GameTime").Value) <  CDATE(lineupTimeChk)) and (CDATE(selectDate) = date()))then %>
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><blackIcon><i class="fas fa-user-lock"></i></blackIcon></td>									
										<%elseif ((objsrAvailPlayers.Fields("PID").Value = startingForwardPID) or (objsrAvailPlayers.Fields("PID").Value = startingForward2PID)) then%>
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" checked id="C" name="sForward" value="<%=objsrAvailPlayers.Fields("firstName").Value & " " & objsrAvailPlayers.Fields("lastName").Value & ";" & objsrAvailPlayers.Fields("PID").Value & ";" & objsrAvailPlayers.Fields("barps").Value  & ";" &  objsrAvailPlayers.Fields("GameTime").Value & ";"%>"></td>
										<%else%>
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" id="C" name="sForward" value="<%=objsrAvailPlayers.Fields("firstName").Value & " " & objsrAvailPlayers.Fields("lastName").Value & ";" & objsrAvailPlayers.Fields("PID").Value & ";" & objsrAvailPlayers.Fields("barps").Value  & ";" &  objsrAvailPlayers.Fields("GameTime").Value & ";"%>"></td>
										<%end if%>
									</tr>									
									<%
										objsrAvailPlayers.MoveNext
										Wend
										
									 objsrAvailPlayers.close
									 loopCt = loopCt + 1
								   Wend										
									%>
							</table>
							<br>
						</div>
					</div>
					<% else %>
					<div class="row">
						<div class="col-xs-12">
							<div class="jumbotron" style="background-color:#9a1400;color:white;text-align:center;vertical-align:middle">
								<i class="fas fa-frown big"></i>&nbsp;No Forwards Available Tonight!
							</div>
						</div>
					</div>
					<% end if%> 	
					
					<%
	   				objsrAvailPlayers.Open      "SELECT * FROM qLineupDaily WHERE (POS = 'GUA' Or POS = 'G-F') and qLineupDaily.OwnerID  = "&ownerid&" and gameday = cdate('"&selectDate&"')",  objConn,1,1
						guardCnt = objsrAvailPlayers.RecordCount
						objsrAvailPlayers.CLOSE
					%>
					<%if guardCnt > 0 Then%>
					<div class="row">
						<div class="col-xs-12">
							<table class="table table-custom-black table-bordered table-responsive table-condensed">
								<tr>
									<th class="text-uppercase text-left;" colspan="2">GUARDS</th>
									<th style="text-align:center;">AVG</th>
									<th style="text-align:center;">L/5</th>
									<th style="text-align:center;">USG</th>
									<th style="text-align:center;"><i class="fas fa-basketball-hoop"></i></th>
								</tr>
							<%
							
					   loopCt = 1
					   while loopCt <= 2 
							if loopCt = 1 then
											objsrAvailPlayers.Open  "SELECT * FROM qLineupDaily WHERE (POS = 'GUA' Or POS = 'G-F') " &_
					                                    "AND qLineupDaily.OwnerID  = "&ownerid&" AND gameday = cdate('"&selectDate&"') " & _
					  	                                "AND PID in ("&startingGuardPID&","&startingGuard2PID&")",  objConn,1,1
							else
											objsrAvailPlayers.Open  "SELECT * FROM qLineupDaily WHERE (POS = 'GUA' Or POS = 'G-F') " &_
					                                    "AND qLineupDaily.OwnerID  = "&ownerid&" AND gameday = cdate('"&selectDate&"') " & _
					  	                                "AND PID Not in ("&startingGuardPID&","&startingGuard2PID&")",  objConn,1,1
							end if
								 
							While Not objsrAvailPlayers.EOF
							 if len(objsrAvailPlayers.Fields("GameTime").Value) = 10 then
									wtime = Left(objsrAvailPlayers.Fields("GameTime").Value,4) & Right(objsrAvailPlayers.Fields("GameTime").Value,3)
							 else
									wtime = Left(objsrAvailPlayers.Fields("GameTime").Value,5) & Right(objsrAvailPlayers.Fields("GameTime").Value,3)
							 end if
							%>
							<%if ((objsrAvailPlayers.Fields("PID").Value = selectedGuard1) or (objsrAvailPlayers.Fields("PID").Value = selectedGuard2)) then%>
								<tr class="success">
							<%else%>
								<tr style="background-color:white;">
							<%end if%>
								<!--#include virtual="Common/availPlayers.inc"-->							
							<%if g1DeadlinePassed  AND g2DeadlinePassed then%>
								<%if ((objsrAvailPlayers.Fields("PID").Value = selectedGuard1) or (objsrAvailPlayers.Fields("PID").Value = selectedGuard2)) then%>
									<td class="success" style="vertical-align:middle;text-align:center;width:10%"><greenIcon2><i class="fas fa-sync fa-spin"></i></greenIcon2></td>							
								<%else%>
									<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><redIcon2><i class="fas fa-user-slash"></i></redIcon2></td>
								<%end if%>
							<%elseif ((objsrAvailPlayers.Fields("PID").Value = selectedGuard1) or (objsrAvailPlayers.Fields("PID").Value = selectedGuard2))then%>
									<td class="success" style="vertical-align:middle;text-align:center;width:10%"><greenIcon2><i class="fas fa-sync fa-spin"></i></greenIcon2></td>
							<%elseif ((CDATE(objsrAvailPlayers.Fields("GameTime").Value) <  CDATE(lineupTimeChk)) and (CDATE(selectDate) = date()))then%>
									<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><blackIcon><i class="fas fa-user-lock"></i></blackIcon></td>
							<%elseif ((objsrAvailPlayers.Fields("PID").Value = startingGuardPID)  or (objsrAvailPlayers.Fields("PID").Value = startingGuard2PID)) then%>
									<td  style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" checked id="C" name="sGuard" value="<%=objsrAvailPlayers.Fields("firstName").Value & " " & objsrAvailPlayers.Fields("lastName").Value & ";" & objsrAvailPlayers.Fields("PID").Value & ";" &  objsrAvailPlayers.Fields("barps").Value  & ";" &  objsrAvailPlayers.Fields("GameTime").Value & ";"%>"></td>
							<%else%>
									<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" id="C" name="sGuard" value="<%=objsrAvailPlayers.Fields("firstName").Value & " " & objsrAvailPlayers.Fields("lastName").Value & ";" & objsrAvailPlayers.Fields("PID").Value & ";" & objsrAvailPlayers.Fields("barps").Value  & ";" &  objsrAvailPlayers.Fields("GameTime").Value & ";"%>"></td>
							<%end if%>							
							</tr>						
					
						<%
							objsrAvailPlayers.MoveNext
							Wend							
							objsrAvailPlayers.close
							loopCt = loopCt + 1
							Wend										
						%>
						</table>
						</div>
					</div>
					<% else %>
					<div class="row">
						<div class="col-xs-12">
							<div class="jumbotron" style="background-color:#9a1400;color:white;text-align:center;vertical-align:middle">
								<i class="fas fa-frown big"></i>&nbsp;No Guards Available Tonight!
							</div>
						</div>
					</div>
					<% end if 
					objsrAvailPlayers.close
					%>	
					</br>
					<div class="row">
						<div class="col-xs-12">
						<%if cDeadlinePassed = false or f1DeadlinePassed = false or f2DeadlinePassed = false or g1DeadlinePassed = false or g2DeadlinePassed = false then%>
							<button type="submit" id="idSubmitLineup"  value="Submit Lineup" name="Action" class="btn   btn-block btn-default"><span class="glyphicon glyphicon-save"></span>&nbsp;Submit Lineup</button>
						<%end if%>
						</div>			
					</div>			
					</br>
					<% if unAvailCnt > 0 then %>
					<div class="row">
						<div class="col-xs-12">	
							<div class="panel panel-override">
								<table class="table table-custom-black table-bordered table-responsive table-condensed">
									<tr>
										<td style="border-radius: unset;vertical-align:middle;text-align:center;background-color:black;color:white;text-transform:uppercase;font-weight:bold;" colspan="5"><i class="fas fa-power-off"></i>&nbsp;<%= (FormatDateTime(selectDate,1)) %></td>
									</tr>
									<tr>
										<th style="width:54%;" class="text-uppercase text-center">Player</th>
										<th style="text-align:center;width:12%;">AVG</th>
										<th style="text-align:center;width:12%;">L/5</th>
										<th style="text-align:center;width:12%;">USG</th>
										<th style="text-align:center;width:10%;"><i class="fas fa-basketball-hoop"></i></th>
									</tr>
									<%
										While Not objRSUnavail.EOF
									%>
									<tr style="background-color:white;text-align:left;">
										<td class="big" style="text-align:left"><a class="blue" href="playerprofile.asp?pid=<%=objRSUnavail.Fields("PID").Value %>"><%=left(objRSUnavail.Fields("firstName").Value,14)%>&nbsp;<%=left(objRSUnavail.Fields("lastName").Value,14)%></a>
										<br><small><span class="greenTrade text-uppercase"><%=objRSUnavail.Fields("TeamName").Value%></span>&nbsp;<span class="orange"><%=objRSUnavail.Fields("pos").Value%></span></small>
										<% if objRSUnavail.Fields("IR").Value = true then %>
										<strong>	<i class="fas fa-briefcase-medical red"></i></strong>
										<%end if%>
										<% if objRSUnavail.Fields("pendingtrade").Value = true then %>
										<strong>	<i class="far fa-exchange"></i></strong>
										<%end if%>
											<% if objRSUnavail.Fields("pendingWaiver").Value = true then %>
										<strong>	<i class="fas fa-user-clock auctionText"></i></strong>
											<%end if%>									
										</td>					
										<td style="text-align:center;vertical-align:middle;" class="text-uppercase big"><span class="badgeBlue"><%=round(objRSUnavail.Fields("barps").Value,0)%></span></td>		
											<% if CDbl(objRSUnavail.Fields("l5barps").Value) > CDbl(objRSUnavail.Fields("barps").Value) then %>
												<td style="vertical-align:middle;text-align:center"><span class="badgeUp big"><%= round(objRSUnavail.Fields ("l5barps").Value,0) %></span></td>
											<% elseif CDbl(objRSUnavail.Fields("barps").Value) > CDbl(objRSUnavail.Fields("l5barps").Value) then%>
												<td style="vertical-align:middle;text-align:center"><span class="badgeDown big"><%= round(objRSUnavail.Fields ("l5barps").Value,0) %></span></td>
											<%else%>
											<td class="big" style="vertical-align:middle;text-align:center"><span class="badgeEven"><%= round(objRSUnavail.Fields ("l5barps").Value,0) %></span></td>
											<%end if %>	

											<%if objRSUnavail.Fields ("usage").Value > 0 then %>
												<td class="big" style="vertical-align:middle;text-align:center;"class="text-uppercase"><span class="badgeUsage big"><%= round(objRSUnavail.Fields ("usage").Value,0) %></span></td>
											<%else %>
												<td class="big" style="vertical-align:middle;text-align:center;"class="text-uppercase"><span class="badgeUsage big">0</span></td>
											<%end if%>
										</td>
										<td class="big" style="vertical-align:middle;text-align:center;"class="text-uppercase">OFF</td>										
									</tr>
									<%
									objRSUnavail.MoveNext					
									Wend
									%>
								</table>
							</div>
						</div>
					</div>
					<%end if%>
					<%  							   
					objRSWork.Open "SELECT * FROM tblplayers " & _
												 "WHERE ownerid = "&ownerId&" AND " & _
												 "(ir <> 0 or pendingTrade <> 0 or pendingwaiver <> 0 or rentalplayer <> 0) ", objConn,3,3,1
																 
					if objRSWork.RecordCount > 0 then							   
					%>
					<div class="row">
						<div class="col-xs-12">
							<div style="text-align: center;"><strong><mark>Legend:</mark></strong>&nbsp;&nbsp;<i class="fas fa-briefcase-medical red fa-md"></i>&nbsp;<strong>IR</strong>&nbsp;&nbsp;<i class="far fa-exchange fa-md"></i>&nbsp;<strong>Pending Trade</strong>&nbsp;&nbsp;<i class="fas fa-user-clock fa-md auctionText"></i>&nbsp;<strong>Pending Waiver</strong>&nbsp;&nbsp;<i class="fa fa-registered blue fa-md" aria-hidden="true"></i>&nbsp;<strong>Rented</strong></div>
						</div>
					</div>
					</br>
					<% end if 
					objRSWork.Close 
					%>
				</form>	
				<%end if
				if sAction = "Submit Lineup"  and errorCode = "" then %>
				<form action="dashboard.asp" method="POST" name="frmconfirm">
				<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
				<input type="hidden" name="txtOwnerID" value="<%= ownerid %>" />
				<%
					Dim objRSLineupsPics  

					Set objRSLineupsPics = Server.CreateObject("ADODB.RecordSet")

					objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wCenterPID  & " = PID " , objConn,3,3,1
					cFirstName = objRSLineupsPics.Fields("firstName").Value
					cLastName  = objRSLineupsPics.Fields("lastName").Value
					objRSLineupsPics.Close

					objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wForwardPID  & " = PID " , objConn,3,3,1
					f1FirstName= objRSLineupsPics.Fields("firstName").Value
					f1LastName = objRSLineupsPics.Fields("lastName").Value
					objRSLineupsPics.Close

					objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " &  wForward2PID  & " = PID " , objConn,3,3,1
					f2FirstName= objRSLineupsPics.Fields("firstName").Value
					f2LastName = objRSLineupsPics.Fields("lastName").Value
					objRSLineupsPics.Close

					objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wGuardPID  & " = PID " , objConn,3,3,1
					g1FirstName= objRSLineupsPics.Fields("firstName").Value
					g1LastName = objRSLineupsPics.Fields("lastName").Value
					objRSLineupsPics.Close

					objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wGuard2PID & " = PID " , objConn,3,3,1
					g2FirstName= objRSLineupsPics.Fields("firstName").Value
					g2LastName = objRSLineupsPics.Fields("lastName").Value
					objRSLineupsPics.Close 
				%>
				<% sAction = "Return"%>
					<div class="container">
						<div class="row">
							<div class="col-xs-6">
								<select class="form-control " name="gameDays">
									<%if sAction = "Retrieve Lineup" then%>
									<option value="<%=selectDate%>" selected><%=selectDate%>  - Next Game Date</option>
									<% else %>
									<option value="<%=objRSgames.Fields("gameDay")%>" selected><%=objRSgames.Fields("gameDay")%> - Next Game Date</option>
									<% end if %>
									<% While not objRSgames.EOF %>
									<option value="<%=objRSgames("gameDay")%>"><%=objRSgames.Fields("gameDay")%></option>
									<% objRSgames.MoveNext
									Wend 
									%>
								</select>
							</div>
							<div class="col-xs-6 align="right">
								<button type="submit"	value="Retrieve Lineup" name="Action" class="btn btn-default  btn-block"><span class="glyphicon glyphicon-search"></span>&nbsp;Retrieve Lineup</button>
							</div>
						</div>
						<br>
						<div class="row">
							<div class="col-xs-12">
								<div class="panel panel-override">
									<table class="table table-custom-black table-striped table-bordered table-condensed">
										<tr>
										<% if date() = cdate(gamedate) then %>
											<th colspan="2" class="panel-title">Today's&nbsp;<i class="fal fa-calendar-alt"></i>&nbsp;<%= gamedate %></th>
										<%else%>
											<th colspan="2" class="panel-title">Future&nbsp;<i class="fal fa-calendar-alt"></i>&nbsp;<%= gamedate %></th>
										<%end if%>
										</tr>
										<tr bgcolor="#FFFFFF">
											<td class="big" colspan="2" style="text-align:center;"><a class="blue" href="playerprofile.asp?pid=<%=wCenterPID %>"><%=left(cFirstName,11)%>&nbsp;<%=cLastName%></a>&nbsp;<orangePos>CEN</orangePos></td>
										</tr>
										<tr bgcolor="#FFFFFF">
											<td class="big" width="50%" style="text-align:left"><a class="blue" href="playerprofile.asp?pid=<%=wForwardPID %>"><%=left(f1FirstName,11)%>&nbsp;<%=f1LastName%></a>&nbsp;<orangePos>FOR</orangePos></td>
											<td class="big" width="50%" style="text-align:right"><a class="blue" href="playerprofile.asp?pid=<%=wForward2PID %>"><%=left(f2FirstName,11)%>&nbsp;<%=f2LastName%></a>&nbsp;<orangePos>FOR</orangePos></td>
										</tr>
										<tr bgcolor="#FFFFFF">
											<td class="big" width="50%" style="text-align:left"><a class="blue" href="playerprofile.asp?pid=<%=wGuardPID %>"><%=left(g1FirstName,11)%>&nbsp;<%=g1LastName%></a>&nbsp;<orangePos>GUA</orangePos></td>
											<td class="big" width="50%" style="text-align:right"><a class="blue" href="playerprofile.asp?pid=<%=wGuard2PID %>"><%=left(g2FirstName,11)%>&nbsp;<%=g2LastName%></a>&nbsp;<orangePos>GUA</orangePos></td>
										</tr>
									</table>
								</div>
							</div>
						</div>
						<div class="row">
							<div class="col-xs-12">
								<span><a href="viewLineups.asp" class="btn btn-block  btn-default" style="min-height: 40px;min-width: 40px;">View All Lineups & Matchups</a></span>
							</div>
						</div>
					</div>
				</form>
				<%end if
				if sAction ="Submit Lineup" and tip_time_error_flag = 1  then %>
				<form action="dashboard.asp" method="POST" name="frmreject" language="JavaScript">
					<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
					<div class="container">
						<div class="row">
							<div class="col-xs-12">
								<div class="alert alert-danger">
								 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
									<strong><i class="fa fa-exclamation-triangle fa-md" aria-hidden="true"></i></strong><br> <%=errorcode%> 
								</div>
								<div>
									<button type="submit" value="" name="" class="btn btn-block btn-default"><i class="fas fa-trash-alt"></i>&nbsp;Return to Line-ups!</button>
								</div>
							</div>
						</div>
					</div>
				</form>
				<% elseif sAction ="Submit Lineup" and errorcodeLU = "Invalid Line-up" then  %>
				<form action="dashboard.asp" method="POST" name="frmreject" language="JavaScript">
					<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
					<div class="container">
						<div class="row">
							<div class="col-xs-12">
								<div class="alert alert-danger">
								 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
									<strong><i class="fa fa-exclamation-triangle fa-md" aria-hidden="true"></i>&nbsp;Line-Up Error!</strong><br><br> <%=errorcode%> 
								</div>
								<div>
								 <a href="#" onClick="history.go(-1); return false;"> <button class="btn btn-block btn-default"><i class="fas fa-trash-alt"></i>&nbsp;Return to Line-ups!</button></a>
								</div>
							</div>
						</div>
					</div>
				</form>
				<%end if %>
				</div>
				<!--END OF HOME TEAM MENU-->

				<!--START OF EDIT TEAM MENU-->
				<div id="team" class="tab-pane fade">
				<!--#include virtual="Common/editteam.inc"-->
				</div>
				<!--END OF EDIT TEAM MENU-->
				
				<!--START OF OTB MENU-->
				<div id="otb" class="tab-pane fade">
				<!--#include virtual="Common/otb.inc"-->
				</div>
				<!--END OF OTB MENU-->
			</div>
		</div>
	</div>
</div>
<%
<!--CLOSE CONNECTIONS-->
	objRS.Close
  objRSteamsneeds.Close
  objRSteams.Close
	objrsForecast.Close	
	objRSNext5.Close
  objRSflex.Close
	objRSUnavail.Close
  objRSteamRec.Close
  objRSAll.Close

  ObjConn.Close
	Set objRSteams   = Nothing
	Set objRSteamRec = Nothing
  Set objConn      = Nothing
%>
</body>
</html>