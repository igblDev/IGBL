<!-- #include file="adovbs.inc" -->
<%

    '##############################################################################################################################
	'SETUP.ASP
	'- This module is run to setup the database at the start of the season. 
	'  Steps for Setup
	'  1. Pull the schedule from data source and Load tblLeagueSetup. Team names in LeagueSetup Names must match TblNbaTeams
	'  2. Run Setup
	'  3. Set the Trade Deadline and Playoff Dates.
	'
	'  Find Invalid Team Names in tblLeagueSetup
	'  select * from tblLeagueSetup x where not exists (select 1 from tblNBATeams y where y.teamName = x.TeamName)
	'  select * from tblLeagueSetup x where not exists (select 1 from tblNBATeams y where y.teamName = x.Opponent)
	'  update tblLeagueSetup set gamedate = DateAdd("yyyy", 1, gamedate) where gameDate < now()
	'##############################################################################################################################

	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""

	Dim objConn, w_action, teamNum
	Dim objRSteam
	Dim ATLGrid,BKNGrid,BOSGrid,CHAGrid,CHIGrid,CLEGrid,DALGrid,DENGrid,DETGrid,GSWGrid,HOUGrid,INDGrid,LACGrid,LALGrid,MEMGrid,MIAGrid,MILGrid,MINGrid
	Dim NOPGrid,NYKGrid,OKCGrid,ORLGrid,PHIGrid,PHXGrid,PORGrid,SACGrid,SASGrid,TORGrid,UTAGrid,WASGrid
	

  Set objConn	= Server.CreateObject("ADODB.Connection")
  Set objRSALLGameDays = Server.CreateObject("ADODB.RecordSet")
  Set objRSOneGame = Server.CreateObject("ADODB.RecordSet")
  Set objRPoints = Server.CreateObject("ADODB.RecordSet")
  
  objConn.Open Application("lineupstest_ConnectionString")
  objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=lineupstest.mdb;" & _
  	              "Persist Security Info=False"		  

	Dim objRSwaivers, w_priority, I
	Set objRSwaivers 	= Server.CreateObject("ADODB.RecordSet")

	w_action = Request.Form("action")
	'Response.Write "Value is: " & w_Action & ".<br>"

  If Request.Form("action") = "Save Form Data" Then

	 '########### Delete all rows from IGBL_SCHEDULE ###########	  
	 strSQL = "DELETE FROM TBL_AUCTION_RECAP"
	 objConn.Execute strSQL
	 Response.Write "Sql = " & strSQL  & ".<br>"
	 
	 '########### Delete all rows from IGBL_SCHEDULE ###########	  
	 strSQL = "DELETE FROM TBL_SCHEDULE"
	 objConn.Execute strSQL
	 Response.Write "Sql = " & strSQL  & ".<br>"
	 
	 '########### Insert rows to IGBL_SCHEDULE ###########	  
	 strSQL = "INSERT INTO TBL_SCHEDULE SELECT * FROM TBL_SCHEDULE_ORIG"
	 objConn.Execute strSQL
	 Response.Write "Sql = " & strSQL  & ".<br>"

	 '########### Delete all rows from tblGameDeadLines ###########
	 strSQL = "delete from tblGameDeadLines "
	 objConn.Execute strSQL
	 Response.Write "Sql = " & strSQL  & ".<br>"
	  
     '########### Delete all rows from NBAINDTMSked ###########	  
	 strSQL = "delete from NBAINDTMSked "
	 objConn.Execute strSQL
	 Response.Write "Sql = " & strSQL  & ".<br>"
	 
	 '########### Delete all rows from tblGameGrid ###########	  
	 strSQL = "delete from tblGameGrid "
	 objConn.Execute strSQL
	 Response.Write "Sql = " & strSQL  & ".<br>"
	 
	 strSQL = "DELETE FROM tbl_points_scored"
     objConn.Execute strSQL
	 Response.Write "Sql = " & strSQL  & ".<br>"
	
	 strSQL = "DELETE from tblTransactions"
 	 objConn.Execute strSQL
	 Response.Write " " & strSQL  & " <br>"

	strSQL = "DELETE from tblPendingTrades"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from Standings_Cycle"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"

	strSQL = "DELETE from tblTrades"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"

   	strSQL = "DELETE from tblWaivers"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"

	strSQL = "DELETE from tblwaiverlog"
	 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"

   	strSQL = "DELETE from tbl_Lineups"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from tbl_lineups_staggered"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from tbl_lineups_history"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"	
	
	strSQL = "DELETE from newbox"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from tblTradeAnalysis" 
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from tbl_barps" 
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from tblLast5" 
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from Standings_RD1" 
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from Standings_RD2" 
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	strSQL = "DELETE from Standings_RD3" 
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
		
	strSQL = "update Standings " &_
             "SET rank=1,won=0,loss=0,PPG=null,OPPG=null,DIFF=null,GB=null,ST=null,GR=null,JP=null,DB=null,CJ=null,JW=null,MJ=null,AW=null," & _
			 "FC=null,TA=null,DM=null,LP=null,cycle=null,prs=0"
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"
	
	'****************************************
	'* Initialize Players
	'****************************************
	strSQL = "UPDATE TBLPLAYERS " & _
			 "SET clearwaiverdate = #10/23/2018#, pendingwaiver = 0, ontheblock = 0, pendingtrade = 0, rentalPlayer = 0, " & _
			 "LastTeamInd = null, ownerID = 0, playerStatus = 'W', l5barps = 0, auctionPrice = 0, ir = 0, injury = 0 "
 	objConn.Execute strSQL
	Response.Write " " & strSQL  & " <br>"

	'****************************************
	'* SET Stagger_window and Timed Events
	'****************************************
	strSQL = "update tblParameterCtl set param_indicator = 0 where param_name = 'STAGGER_WINDOW' "
	objConn.Execute strSQL
			
	strSQL = "update tbltimedEvents SET nextrun_EST = '10/23/2018 1:00:00 PM' WHERE event = 'pendingwaiversall' "
	objConn.Execute strSQL
	
	strSQL = "update tbltimedEvents SET nextrun_EST = '10/23/2018 7:00:00 PM' WHERE event = 'setwaiversall' "
	objConn.Execute strSQL
	
	strSQL = "update tbltimedEvents SET nextrun_EST = '10/24/2018 7:00:00 PM' WHERE event = 'setStaggeredAll' "
	objConn.Execute strSQL
		
	 
	'*******************************************
	'* SET All Owner Level Options to Yes
	'* Set ActivePlayerCnt = 0, Waiverbal = 250
	'******************************************* 
	strSQL = "update tblOwners " & _
	         "SET receiveFreeAgentAlerts = 1, receiveTradeAlerts = 1, receiveWaiverAlerts = 1, receiveStaggerAlerts = 1, " & _
	         "receiveBoxScoreAlerts = 1, receiveRentalAlerts = 1, receiveOTBAlerts = 1, receiveEmails = 1, acceptTradeOffers = 1, " &_
             "receiveEmailLeagueAlerts = 1, receiveTexts = 1, seasonOver = 0, ontheblockcomments=null, porank=null,   " & _
		     "ontheblockneedscen=0, ontheblockneedsfor=0, ontheblockneedsfc=0, ontheblockneedsgua=0, ontheblockneedsgf=0, ontheblockall=0, " &_
		     "ActivePlayerCnt=0, waiverbal=250 " &_			 
		     "Where ownerid <> 99 "
	objConn.Execute strSQL
	
	  '********************************************************
	  'Update waiver priority.
	  '********************************************************
	  strSQL ="update TBLOWNERS set WaiverPriority = 1 where ownerid = 1 "  'Gary
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 2 where ownerid = 2 "  'Family Guys
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 3 where ownerid = 3 "  'Babineaux
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 4 where ownerid = 4 "  'CJ
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 5 where ownerid = 5 "  'Jack
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 6 where ownerid = 6 "  'LaMont
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 7 where ownerid = 7 "  'Banks
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 8 where ownerid = 8 "  'Fred
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 9 where ownerid = 9 "  'Denver Guys
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	  
	  strSQL ="update TBLOWNERS set WaiverPriority = 10 where ownerid = 10 "  'Dee
	  objConn.Execute strSQL
	  Response.Write "Sql = " & strSQL  & "<br>"
	 
	 
	 '********************************************************
	 'LeagueSetup Names must match TblNbaTeams
	 '********************************************************
     objRSALLGameDays.Open "select count(*) as TotalTeams, gamedate, min(TipTimeEST - 1/24) as EarlyTip, max(TipTimeEst - 1/24) as LateTip " & _
                        "from tblLeagueSetup where gamedate >= #10/24/2018# " & _
                        "group by gamedate " & _
                        "having count(*) >= 14 " & _
                        "order by gamedate", objConn
	  loopCtl = 1
	  wCycle = 1
	  While loopCtl <= 91 	
	     wGameDay = objRSALLGameDays.Fields("gamedate").Value
		 wNumTeams = objRSALLGameDays.Fields("TotalTeams").Value
		 wEarlyTip = objRSALLGameDays.Fields("EarlyTip").Value
		 wLateTip = objRSALLGameDays.Fields("LateTip").Value
		 	 
		 
		 if loopCtl = 10 OR loopCtl = 20 OR loopCtl = 30 OR loopCtl = 40 OR loopCtl = 50 OR loopCtl = 60 OR loopCtl = 70 then
		    wGameType = "PRG"
	     elseif loopCtl = 11 OR loopCtl = 21 OR loopCtl = 31 OR loopCtl = 41 OR loopCtl = 51 OR loopCtl = 61 then
		    wGameType = "RSG"
			wCycle = wCycle + 1
		 elseif loopCtl > 70 then
		    wGameType = "PSG"
			wCycle = 99
	     else
		    wGameType = "RSG"
		 end if
		 
	     Response.Write "Game "&loopCtl&" - "&wGameDay& ". Teams = "&wNumTeams&".  GameType = "&wGameType&". Cycle = "&wCycle&". <br>"
		 
		 '########### Insert to tblGameDeadLines ###########
		 strSQL = "insert into tblGameDeadLines (GameDay, GameDeadline, gameStaggerDeadline, gameType, IgblGameNum, cycle) " & _
		           "values (#"&wGameDay&"#,'"&wEarlyTip&"','"&wLateTip&"','"&wGameType&"',"&loopCtl&","&wCycle&")"
					 
         objConn.Execute strSQL				  				  					 
		 'Response.Write "*****strSQL = "&strSQL&"<br>"
		 
		 objRSOneGame.Open "SELECT t.NbaTid, s.TeamName, s.GameDate, s.TipTimeEst - 1/24 as TipTime, s.GameLoc, s.Opponent, " & _
		                           "t2.teamShortName " & _
                           "FROM tblleaguesetup s, tblNBATeams t, tblNBATeams t2 " & _
                           "WHERE s.TeamName = t.teamName " & _
                           "AND s.Opponent = t2.teamName " & _
                           "AND s.gamedate = #"&wGameDay&"#", objConn,3,3,1
						   
		 FuncCall = InitializeGridValues()	
		 While Not objRSOneGame.EOF
			wTeamID = objRSOneGame.Fields("NbaTid").Value
			wGameTime = objRSOneGame.Fields("TipTime").Value
			if objRSOneGame.Fields("GameLoc").Value = "at" OR objRSOneGame.Fields("GameLoc").Value = "@"   then
				wOpponent = "at "&objRSOneGame.Fields("teamShortName").Value
				wLongOpp  = "at "&objRSOneGame.Fields("Opponent").Value
			else
				wOpponent = "vs "&objRSOneGame.Fields("teamShortName").Value		    
				wLongOpp  = "vs "&objRSOneGame.Fields("Opponent").Value
			end if
		 
	        '########### Insert to NBAINDTMSked ###########
			strSQL = "insert into NBAINDTMSked (NBATeam, GameDay, GameTime, OppLongName, Opponent) " & _
		             "values ("&wTeamID&",#"&wGameDay&"#,'"&wGameTime&"','"&wLongOpp&"','"&wOpponent&"')"
					 
            objConn.Execute strSQL				  				  
			'Response.Write "*****strSQL = "&strSQL&"<br>"
			
			FuncCall = UpdateGridField()
			objRSOneGame.MoveNext
		 Wend		 
		 objRSOneGame.Close
	      
		 '########### Insert to tblGameGrid ###########
         strSQL = "insert into tblGameGrid " & _
		          "values (#"&wGameDay&"#,"&ATLGrid&","&BKNGrid&","&BOSGrid&","&CHAGrid&","&CHIGrid&","&CLEGrid&","&DALGrid&","&DENGrid&","&DETGrid&","&GSWGrid&"," _
				                           &HOUGrid&","&INDGrid&","&LACGrid&","&LALGrid&","&MEMGrid&","&MIAGrid&","&MILGrid&","&MINGrid&","&NOPGrid&","&NYKGrid&"," _
										   &OKCGrid&","&ORLGrid&","&PHIGrid&","&PHXGrid&","&PORGrid&","&SACGrid&","&SASGrid&","&TORGrid&","&UTAGrid&","&WASGrid&")"
										   
         objConn.Execute strSQL				  				  										   
         'Response.Write "strSQL = "&strSQL&"<br>"
		 
		 
		 '########### Update tbl_schedule ###########		 
		 strSQL = "update tbl_schedule set GameDay = #"&wGameDay&"# where IgblGameNum = "&loopCtl
		 objConn.Execute strSQL
		  
		 objRSALLGameDays.MoveNext		 		 
		 loopCtl = loopCtl + 1
	  wend	  
	  objRSALLGameDays.Close
	  
	  'FuncCall = BuildPointsScoredTbl()
	  
	  Response.Write "loopCtl = " & loopCtl  & ".<br>"
	   	  		  		
  End if

  Function InitializeGridValues ()	
  
	ATLGrid = "null"
	BKNGrid = "null"
	BOSGrid = "null"
	CHAGrid = "null"
    CHIGrid = "null"
    CLEGrid = "null"
    DALGrid = "null"
    DENGrid = "null"
    DETGrid = "null"
    GSWGrid = "null"	
    HOUGrid = "null"
    INDGrid = "null"
    LACGrid = "null"
    LALGrid = "null"
    MEMGrid = "null"
    MIAGrid = "null"
    MILGrid = "null"
    MINGrid = "null"
    NOPGrid = "null"
    NYKGrid = "null"
    OKCGrid = "null"
    ORLGrid = "null"
    PHIGrid = "null"
    PHXGrid = "null"
    PORGrid = "null"
    SACGrid = "null"
    SASGrid = "null"
    TORGrid = "null"
    UTAGrid = "null"
    WASGrid = "null"	
	
  End Function
  
  Function UpdateGridField ()
	
    if wTeamID = 1 then
	   ATLGrid = 1
	elseif wTeamID = 2 then
	   BOSGrid = 1	
	elseif wTeamID = 3 then
	   BKNGrid = 1						
	elseif wTeamID = 4 then
	   CHAGrid = 1		
	elseif wTeamID = 5 then
	   CHIGrid = 1		
	elseif wTeamID = 6 then
	   CLEGrid = 1		
	elseif wTeamID = 7 then
	   DALGrid = 1		
	elseif wTeamID = 8 then
	   DENGrid = 1		
	elseif wTeamID = 9 then
	   DETGrid = 1		
	elseif wTeamID = 10 then
	   GSWGrid = 1		
	elseif wTeamID = 11 then
	   HOUGrid = 1		
	elseif wTeamID = 12 then
	   INDGrid = 1		
	elseif wTeamID = 13 then
	   LACGrid = 1		
	elseif wTeamID = 14 then
	   LALGrid = 1		
	elseif wTeamID = 15 then
	   MEMGrid = 1		
	elseif wTeamID = 16 then
	   MIAGrid = 1		
	elseif wTeamID = 17 then
	   MILGrid = 1		
	elseif wTeamID = 18 then
	   MINGrid = 1		
	elseif wTeamID = 19 then
	   NOPGrid = 1		
	elseif wTeamID = 20 then
	   NYKGrid = 1		
	elseif wTeamID = 21 then
	   OKCGrid = 1		
	elseif wTeamID = 22 then
	   ORLGrid = 1		
	elseif wTeamID = 23 then
	   PHIGrid = 1		
	elseif wTeamID = 24 then
	   PHXGrid = 1		
	elseif wTeamID = 25 then
	   PORGrid = 1		
	elseif wTeamID = 26 then
	   SACGrid = 1		
	elseif wTeamID = 27 then
	   SASGrid = 1		
	elseif wTeamID = 28 then
	   TORGrid = 1		
	elseif wTeamID = 29 then
	   UTAGrid = 1		
	elseif wTeamID = 30 then
	   WASGrid = 1			   
	else
	   Response.Write "******** ERROR, ERROR, ERROR, ERROR, ERROR..  wTeamID value is invalid.  Value = "&wTeamID
	end if
				
  End Function

  Function BuildPointsScoredTbl ()	
  
    strSQL = "DELETE FROM tbl_points_scored"
    objConn.Execute strSQL
	
	objRPoints.Open "select * from tbl_schedule order by gameday", objConn,3,3,1
	While Not objRPoints.EOF
	   wGameDay  = objRPoints.Fields("GameDay").Value  
	   wHomeID   = objRPoints.Fields("HomeTeamInd").Value   
	   wAwayID   = objRPoints.Fields("AwayTeamInd").Value   	   
	   wGameNum  = objRPoints.Fields("IGBLGameNum").Value   
	   
	   if IsNull(wHomeID) then
	      x = 1    'Do Nothing
	   else
	      if wHomeID <> 99 then
	         strSQL = "insert into tbl_points_scored (IGBLGameNum,GameDate,HomeTeamId,AwayTeamID) values ("&wGameNum&",#"&wGameDay&"#,"&wHomeID&","&wAwayID&")"
             objConn.Execute strSQL
             Response.Write "Sql = " & strSQL  & ".<br>"   
		  end if
     
	      if wAwayID <> 99 then
		     strSQL = "insert into tbl_points_scored (IGBLGameNum,GameDate,HomeTeamId,AwayTeamID) values ("&wGameNum&",#"&wGameDay&"#,"&wAwayID&","&wHomeID&")"
             objConn.Execute strSQL
             Response.Write "Sql = " & strSQL  & ".<br>"
	      end if
	   end if
	  
	   objRPoints.MoveNext			
	Wend		 
	objRPoints.Close
	
  End Function  
  
  
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>IGBL - Setup</title>
<link rel="stylesheet" type="text/css" href="sample.css">
</head>

<body bgcolor="#FFFFF7">

<h2>Setup (4)</h2>
<p>&nbsp;</p>
<form action="Setup.asp" method="POST">
  <input type="hidden" name="action" value="Save Form Data">

  </div>
    </center>
  </div>

  <br>
  </p>

  <div align="center">
      <center>

  <table width="450" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">
    <tr align="center">
      <td width="100%" align="center"> <input type="submit" value="Run Setup" name="Submit"> </td>
    </tr>
  </table>

      </center>
  </div>

</form>
<%
  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>

</body>

</html>