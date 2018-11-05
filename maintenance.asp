<!-- #include file="adovbs.inc" -->

<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""

	Dim objConn, w_action, teamNum, objrsTest
	Dim objRSteam

  Set objConn	= Server.CreateObject("ADODB.Connection")
  Set objConnStage	= Server.CreateObject("ADODB.Connection")
  Set objrsTest = Server.CreateObject("ADODB.RecordSet")
  Set objRStteam = Server.CreateObject("ADODB.RecordSet")
  Set objRSOwners  = Server.CreateObject("ADODB.RecordSet")
  Set objRSPlayers = Server.CreateObject("ADODB.RecordSet")
  Set objRSAll = Server.CreateObject("ADODB.RecordSet")
  Set objTxnAmt = Server.CreateObject("ADODB.RecordSet")
  Set objRSWork = Server.CreateObject("ADODB.RecordSet")
  Set objRSPower = Server.CreateObject("ADODB.RecordSet")
  Set objRSPoints = Server.CreateObject("ADODB.RecordSet")
  Set objrsWins = Server.CreateObject("ADODB.RecordSet")
  Set objrsTies = Server.CreateObject("ADODB.RecordSet")
  
   
  
   objConn.Open Application("lineupstest_ConnectionString")
   objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
				   
    objConnStage.Open Application("igblDev_ConnectionString")
	objConnStage.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
											"Data Source=igblDev.mdb;" & _
											"Persist Security Info=False"   

	Dim objRSwaivers, w_priority
	Set objRSwaivers 	= Server.CreateObject("ADODB.RecordSet")

	w_action = Request.Form("action")
	'Response.Write "Value is: " & w_Action & ".<br>"

   
  If Request.Form("action") = "Save Form Data" Then
  

  Response.Write "The following Statements have been executed. <br> <br>"
    
	
		 'strSQL = "delete from tblwaivers"
	   'objConn.Execute strSQL
     'esponse.Write "Sql = " & strSQL  & ".<br>"
		 
		 'strSQL = "insert into tbl_lineups_staggered select * from tbl_lineups where gameDay = #10/04/2018#"
	     'objConn.Execute strSQL
         'Response.Write "Sql = " & strSQL  & ".<br>"
		 
		 'strSQL = "delete from standings_cycle"
	     'objConn.Execute strSQL
         'Response.Write "Sql = " & strSQL  & ".<br>"
		 
		 
   
	
      
  '********************************************************
  'Update OWNERS table.
  '********************************************************
    'strSQL = "update tblOwners " & _
	'         "SET receiveFreeAgentAlerts = 0, receiveTradeAlerts = 0, receiveWaiverAlerts = 0, receiveStaggerAlerts = 0, " & _
	'		     "receiveBoxScoreAlerts = 0, receiveRentalAlerts = 0, receiveOTBAlerts = 0, acceptTradeOffers = 1 " & _
	'	     "Where ownerid not in (99,8,10) "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>" 
	
	'strSQL = "update tblOwners SET receiveEmails = 0, receiveTexts = 0 where ownerid not in (99,8,10) "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>" 
	
	
	'strSQL = "update tblowners SET seasonOver = 1 where ownerid in (8)"
 	'objConn.Execute strSQL
	'Response.Write " " & strSQL  & " <br>"
     
  '********************************************************
  'Update Lineups table.
  '********************************************************
  'strSQL ="update tbl_lineups set penalty = 1 where ownerid = 2 and gameday = #8/19/2017# "  
  'objConn.Execute strSQL
  'Response.Write "Sql = " & strSQL  & ".<br>"
  
  'strSQL ="update tbl_lineups_staggered set penalty = 0 where penalty = 1 "  
  'objConn.Execute strSQL
  'Response.Write "Sql = " & strSQL  & ".<br>"
 
   '********************************************************
   'Update waiver priority.
   '********************************************************
   'strSQL ="update TBLOWNERS set WaiverPriority = 1 where ownerid = 1 "  'Gary
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
   
   'strSQL ="update TBLOWNERS set WaiverPriority = 2 where ownerid = 2 "  'Family Guys
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 6 where ownerid = 3 "  'BAbs
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 9 where ownerid = 4 "  'CJ
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 5 where ownerid = 5 "  'Jack
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 4 where ownerid = 6 "  'Matt 
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 10 where ownerid = 7 "  'Banks
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 7 where ownerid = 8 "  'Fred
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 8 where ownerid = 9 "  'Denver Guys
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"
  
   'strSQL ="update TBLOWNERS set WaiverPriority = 3 where ownerid = 10 "  'Dee
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & "<br>"   
   
  '****************************************************************************************************
  'Update NBAINDTMSKed table.
  'When game times change or Postponed games are rescheduled.  Insert to NBAINDTMSked and tblGameGrid
  '****************************************************************************************************
   'strSQL = "update tblGameGrid set DET = null, POR = null where gameDate = #1/7/2017# "
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & ".<br>" 					 
   
   'strSQL = "update tblGameGrid set DET = 1, POR = 1 where gameDate = #1/8/2017# "
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & ".<br>"
   
   'strSQL = "update tblLeagueSetup set gameDate = #1/30/2017# where id in (8094,8097)"
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & ".<br>"
	
  '********************************************************
  'Update PLAYERS table.
  '********************************************************

   'strSQL = "update tblPlayers set lastName = 'Mason III' where pid = 501"
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & ".<br>"
    
   strSQL = "delete from tblwaivers where waiverid = 55"
   objConn.Execute strSQL
   Response.Write "Sql = " & strSQL  & ".<br>"
	 
   'strSQL = "update tblPlayers " & _
   '	       "SET playerStatus = 'F', OwnerId = 0, clearwaiverdate = null, LastTeamInd = null " & _
   '         "WHERE clearwaiverdate <= date() + 1 and playerStatus = 'W'"
   
   'strSQL = "update tblPlayers " & _
   '	        "SET playerStatus = 'F', OwnerId = 0, clearwaiverdate = null, LastTeamInd = null " & _
   '         "WHERE playerStatus = 'W'"
  
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & ".<br>"
   
    'strSQL = "update tblPlayers " & _
    '	     "SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() " & _
    '	     "WHERE playerStatus = 'F' or playerStatus = 'S' or RentalPlayer = 1 "
					
   'objConn.Execute strSQL
   'Response.Write "Sql = " & strSQL  & ".<br>"
   
 
  
   '********************************************************
   'Update Barps table.
   '********************************************************		
    'strSQL = "update tbl_barps SET team = 'NOR' " & _
    '           "WHERE first = 'Toney' and last = 'Douglas' "
    'objConn.Execute strSQL
    'Response.Write "Sql = " & strSQL  & ".<br>"
	
  	 	
	
     'strSQL = "update tblTransactions set TransDate = TransDate + 1 "
	 'objConn.Execute strSQL
	 'Response.Write "Sql = " & strSQL  & ".<br>"
	
    '********************************************************
    'Miscellaneous work
	'Create some test dates for the system
	'Use Set Starting Date in most cases.
    '********************************************************		  	 	 	   	  
	  
	  'objRSWork.Open "SELECT max(gameDay+1) as maxDate FROM tblGameDeadLines", objConn,1,1
	  'if IsNull(objRSWork.Fields("maxDate")) OR (date() > objRSWork.Fields("maxDate")) then
	  '   wNewDate = date()
	  'else
	  '   wNewDate = objRSWork.Fields("maxDate").Value
	  'end if
	  
	  'objRSWork.Close
	  
	  '################################
	  'Get Set of dates to update
	  '################################
	  'objRSWork.Open "SELECT * FROM tblGameDeadLines where gameDay < #8-10-2018# and gameDay < date()  order by gameDay", objConn,1,1
	'  objRSWork.Open "SELECT * FROM tblGameDeadLines where gameDay > date()  order by gameDay", objConn,1,1
	  

	  
	  'Set Starting Date
	  'wNewDate = date()
	 ' wNewDate = #9/24/2018#	
	  
	 ' Response.Write "Start Date = " & wNewDate  & ".<br>"
	  
	  '#########################################################
	  'Use 1 of the while statements.   
	  'Counter is the number of test games days you want to set
	  '#########################################################
	 ' lnCount = 0
	  
	  'while not objrsWork.EOF	  
	  'while lnCount < 91 and not objrsWork.EOF
	'  while lnCount < 21 and not objrsWork.EOF
	'     lnCount = lnCount + 1	 
		 
	'	 wOldDate = objRSWork.Fields("gameDay").Value
	'	 Response.Write "**** Update date "&wOldDate&" to "&wNewDate&"<br>"		 
	'	 		 		 
	'	 strSQL = "update NBAINDTMSKed set Gameday = #"&wNewDate&"# where Gameday = #"&wOldDate&"#"
	'     objConn.Execute strSQL
    '     Response.Write "Sql = " & strSQL  & ".<br>"
	 
	'     strSQL = "update tblGameDeadLines set Gameday = #"&wNewDate&"# where Gameday = #"&wOldDate&"#"
	'     objConn.Execute strSQL
    '     Response.Write "Sql = " & strSQL  & ".<br>"

	   '  strSQL = "update tbl_schedule set Gameday = #"&wNewDate&"# where Gameday = #"&wOldDate&"#"
	  '   objConn.Execute strSQL
     '    Response.Write "Sql = " & strSQL  & ".<br>"
		  
	'	 wNewDate = wNewDate + 1
    '	 objRSWork.MoveNext
	 ' wend
	  
	 ' objRSWork.Close
	 ' Response.Write "lncount = "&lnCount&"<br>"
			 		  
    '********************************************************
	'Create some test dates END
    '********************************************************	
	'email_message = "J. Winslow Added by Fire <br>M. Morros Dropped by Fire<br>J. Winslow Rejected by Titan"
	'Response.Write email_message
	'email_message = replace(replace(replace(email_message, "Added", "A"), "Dropped", "D"), "Rejected", "R")
	'Response.Write "<br><br>"&email_message
	
					  
    '#########################################
    ' Cleanup Opponent Name in tblLeagueSetup
	'#########################################
     
	'strSQL = "Update tblLeagueSetup set Opponent = 'Atlanta Hawks' where Opponent = 'Atlanta' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Boston Celtics' where Opponent = 'Boston' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Brooklyn Nets' where Opponent = 'Brooklyn' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Charlotte Hornets' where Opponent = 'Charlotte' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Chicago Bulls' where Opponent = 'Chicago' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Cleveland Cavaliers' where Opponent = 'Cleveland' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Dallas Mavericks' where Opponent = 'Dallas' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Denver Nuggets' where Opponent = 'Denver' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Detroit Pistons' where Opponent = 'Detroit' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Golden State Warriors' where Opponent = 'Golden State' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Houston Rockets' where Opponent = 'Houston' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Indiana Pacers' where Opponent = 'Indiana' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Los Angeles Clippers' where Opponent in ('LA', 'L.A. Clippers') "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Los Angeles Lakers' where Opponent in ('Los Angeles', 'L.A. Lakers') "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Memphis Grizzlies' where Opponent = 'Memphis' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Miami Heat' where Opponent = 'Miami' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Milwaukee Bucks' where Opponent = 'Milwaukee' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Minnesota Timberwolves' where Opponent = 'Minnesota' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'New Orleans Pelicans' where Opponent = 'New Orleans' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'New York Knicks' where Opponent = 'New York' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Oklahoma City Thunder' where Opponent = 'Oklahoma City' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Orlando Magic' where Opponent = 'Orlando' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Philadelphia 76ers' where Opponent = 'Philadelphia' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Phoenix Suns' where Opponent = 'Phoenix' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Portland Trail Blazers' where Opponent = 'Portland' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Sacramento Kings' where Opponent = 'Sacramento' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'San Antonio Spurs' where Opponent = 'San Antonio' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Toronto Raptors' where Opponent = 'Toronto' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Utah Jazz' where Opponent = 'Utah' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

	'strSQL = "Update tblLeagueSetup set Opponent = 'Washington Wizards' where Opponent = 'Washington' "
	'objConn.Execute strSQL
	'Response.Write "Sql = " & strSQL  & ".<br>"

						
  End if


%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>IGBL - Maintenance</title>
<link rel="stylesheet" type="text/css" href="sample.css">
</head>

<body bgcolor="#FFFFF7">

<h2>Maintenance (4)</h2>
<p>&nbsp;</p>
<form action="maintenance.asp" method="POST">
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
      <td width="100%" align="center"> <input type="submit" value="Run Maintenance" name="Submit"> </td>
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