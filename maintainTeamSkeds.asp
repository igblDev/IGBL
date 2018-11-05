<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%

On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn,objRSgames,objEmail,objLineups,objPlayer,objWork,SsaStagger,Name,objRSTemp,objRSNOTMSkeds,objRSDelete
	Dim strSQL,I,ownerid,sAction
	Dim errorCode, errorDesc, txnteamname, mailTxt
	GetAnyParameter "Action", sAction

	Set objConn       = Server.CreateObject("ADODB.Connection")
	Set objRSTMSkeds  = Server.CreateObject("ADODB.RecordSet")
	Set objRSNOTMSkeds= Server.CreateObject("ADODB.RecordSet")
	Set objRSgames    = Server.CreateObject("ADODB.RecordSet")
	Set objEmail      = Server.CreateObject("ADODB.RecordSet")
	Set objLineups    = Server.CreateObject("ADODB.RecordSet")
	Set objPlayer     = Server.CreateObject("ADODB.RecordSet")
	Set objWork       = Server.CreateObject("ADODB.RecordSet")
	Set SsaStagger    = Server.CreateObject("ADODB.RecordSet")
	Set objRSDelete   = Server.CreateObject("ADODB.RecordSet")
	


	objConn.Open Application("lineupstest_ConnectionString")


  objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"

%>
	<!--#include virtual="Common/session.inc"-->
<%
  GetAnyParameter "Action", sAction
	GetAnyParameter "var_sPid", ppPID	
	lineupTimeChk = time() - 1/24	
		select case sAction
	
	  case ""
			'PROCEED TO SCREEN
	  
		case "Retrieve Teams"
				currentDate = Request.Form("GameDays")	
							
		case "Update Times"
			gameTimeCnt   = Request.Form("chkTeamSked").count
			splitNBATMIDS = Split(Request.Form("chkTeamSked"),";")
			newTipTime    = Request.Form("newTime")
			currentDate   = Request.Form("gameDay")				
																	
				
		   teamList = splitNBATMIDS(0)&splitNBATMIDS(1)
		   strSQL = "UPDATE NBAINDTMSKed set GameTime = '"&newTipTime&"' WHERE gameday = #"&currentDate&"# and NbaTeam in ("&teamList&")"
		   objConn.Execute strSQL
		   
		   iTeamId = splitNBATMIDS(0)
		   objWork.Open "select * from tblNBATeams where NBATID = "&iTeamId&" ", objConn,3,3,1
		   sTeams = objWork.Fields("teamShortName").Value
		   objWork.Close
		   
		   iTeamId = mid(splitNBATMIDS(1), 3)
		   objWork.Open "select * from tblNBATeams where NBATID = "&iTeamId&" ", objConn,3,3,1
		   sTeams = sTeams&", "&objWork.Fields("teamShortName").Value
		   objWork.Close
		   
		   Response.Write "Game Day = "&currentDate&"<br>"
		   Response.Write "New Time = "&newTipTime&"<br>"
		   Response.Write "NBAINDTMSKed updated for: "&sTeams&"<br>"
					 
		   
		   'Check if Stagger Window needs to open earlier.
		   if cDate(currentDate) = date() then
				SsaStagger.Open "SELECT * FROM tbltimedEvents where event = 'setStaggeredAll'",objConn,3,3,1
				sTaggerTime = SsaStagger.Fields("nextrun_EST").Value
			    Response.Write "sTaggerTime = "&sTaggerTime&"<br>"
				SsaStagger.Close			
		 
				compareTime = date() + cdate(newTipTime) + 1/24
				Response.Write "compareTime = "&compareTime&"<br>"
		 
				'Do not update if the stagger time has already run for today.
			    if (cdate(compareTime) < cdate(sTaggerTime)) and (cdate(sTaggerTime) < (date() + 1) ) then
					strSQL = "update tbltimedEvents SET nextrun_EST = '"&compareTime&"' WHERE event = 'setStaggeredAll' "				     
				    objConn.Execute strSQL
				    Response.Write "setStaggeredAll updated to = "&compareTime&" EST <br>"
			    end if
		   end if
		   
		   objLineups.Open  "select * from tbl_lineups tl " &_
						 "where gameday = #"&currentDate&"# " &_
						 "and exists " &_
						   "(select 1 from tblplayers t " &_
							"where (pid = tl.sCenter   and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sForward  and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sForward2 and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sGuard    and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sGuard2   and NBATEAMID in ("&teamList&") ) )", objConn,3,3,1
		   
		   Response.Write "<br>Updating "&objLineups.RecordCount&" Lineups for "&currentDate&"<br>"	
		   While Not objLineups.eof
		   
			   iownerId = objLineups.Fields("ownerId").Value
			   objWork.Open "select * from tblowners where ownerId = "&iownerId&" ", objConn,3,3,1
			   sOwner = objWork.Fields("ShortName").Value
			   objWork.Close
			   
			  '############################################################
			  ' Check Each Player to see which time needs to be updated.
			  '############################################################
			  
			  '** Center **
			  iPlayer = objLineups.Fields("sCenter").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "UPDATE tbl_lineups set sCenterTip = '"&newTipTime&"' WHERE gameday = #"&currentDate&"# and OwnerId = "&iownerId
				 objConn.Execute strSQL
				 Response.Write "--> Updated center Time for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close
			  
			  '** Forward 1 **
			  iPlayer = objLineups.Fields("sForward").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "UPDATE tbl_lineups set sForwardTip = '"&newTipTime&"' WHERE gameday = #"&currentDate&"# and OwnerId = "&iownerId
				 objConn.Execute strSQL
				 Response.Write "--> Update Forward1 time for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close
			  
			  '** Forward 2 **
			  iPlayer = objLineups.Fields("sForward2").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "UPDATE tbl_lineups set sForwardTip2 = '"&newTipTime&"' WHERE gameday = #"&currentDate&"# and OwnerId = "&iownerId
				 objConn.Execute strSQL
				 Response.Write "--> Update Forward2 time for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close

			  '** Guard 1 **
			  iPlayer = objLineups.Fields("sGuard").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "UPDATE tbl_lineups set sGuardTip = '"&newTipTime&"' WHERE gameday = #"&currentDate&"# and OwnerId = "&iownerId
				 objConn.Execute strSQL
				 Response.Write "--> Update Guard1 time for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close

			  '** Guard 2 **
			  iPlayer = objLineups.Fields("sGuard2").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "UPDATE tbl_lineups set sGuardTip2 = '"&newTipTime&"' WHERE gameday = #"&currentDate&"# and OwnerId = "&iownerId
				 objConn.Execute strSQL
				 Response.Write "--> Update Guard2 time for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close				  
			  
			  objLineups.MoveNext
		   Wend
		   objLineups.Close
			
			sAction = ""	 			
			
		case "Delete Teams"
			gameTimeCnt  = Request.Form("chkTeamSked").count
			splitNBATMIDS= Split(Request.Form("chkTeamSked"),";")
			newTipTime   = Request.Form("newTime")
			currentDate  = Request.Form("gameDay")   
			teamList     = splitNBATMIDS(0)&splitNBATMIDS(1)
			c_Time_1159      = "11:59:59 PM"
			
			iTeam1 = splitNBATMIDS(0)
            iTeam2 = mid(splitNBATMIDS(1),3)
			
			Response.Write "iTeam1 = "&iTeam1&"<br>"	
			Response.Write "iTeam2 = "&iTeam2&"<br>"	
			
			strSQL = "DELETE FROM NBAINDTMSKed WHERE gameday = #"&currentDate&"# and NbaTeam in ("&teamList&")"
			Response.Write strSQL
  		    objConn.Execute strSQL
			
			objLineups.Open  "select * from tbl_lineups tl " &_
						 "where gameday = #"&currentDate&"# " &_
						 "and exists " &_
						   "(select 1 from tblplayers t " &_
							"where (pid = tl.sCenter   and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sForward  and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sForward2 and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sGuard    and NBATEAMID in ("&teamList&") ) OR " &_
								  "(pid = tl.sGuard2   and NBATEAMID in ("&teamList&") ) )", objConn,3,3,1
		   
		   Response.Write "<br>Updating "&objLineups.RecordCount&" Lineups for "&currentDate&"<br>"	
		   While Not objLineups.eof
		   
			   iownerId = objLineups.Fields("ownerId").Value
			   objWork.Open "select * from tblowners where ownerId = "&iownerId&" ", objConn,3,3,1
			   sOwner = objWork.Fields("ShortName").Value
			   objWork.Close
			   
			  '############################################################
			  ' Check Each Player to see which players needs to be removed
			  '############################################################
			  
			  '** Center **
			  iPlayer = objLineups.Fields("sCenter").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "update tbl_lineups set sCenter = 9998, sCenterBarps = 0, sCenterTip = '"&c_Time_1159&"' "&_ 
				          "where ownerID = "&iownerId&" and gameday = #"&currentDate&"#"
				 objConn.Execute strSQL
				 Response.Write "--> Removed center for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close
			  
			  '** Forward 1 **
			  iPlayer = objLineups.Fields("sForward").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "update tbl_lineups set sForward = 9996, sForwardBarps = 0, sForwardTip = '"&c_Time_1159&"' "&_ 
				          "where ownerID = "&iownerId&" and gameday = #"&currentDate&"#"
				 objConn.Execute strSQL
				 Response.Write "--> Removed Forward1 for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close
			  
			  '** Forward 2 **
			  iPlayer = objLineups.Fields("sForward2").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "update tbl_lineups set sForward2 = 9997, sForward2Barps = 0, sForwardTip2 = '"&c_Time_1159&"' "&_ 
				          "where ownerID = "&iownerId&" and gameday = #"&currentDate&"#"
				 objConn.Execute strSQL
				 Response.Write "--> Removed Forward2 for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close

			  '** Guard 1 **
			  iPlayer = objLineups.Fields("sGuard").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "update tbl_lineups set sGuard = 9994, sGuardBarps = 0, sGuardTip = '"&c_Time_1159&"' "&_ 
				          "where ownerID = "&iownerId&" and gameday = #"&currentDate&"#"
				 objConn.Execute strSQL
				 Response.Write "--> Removed Guard1 for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close

			  '** Guard 2 **
			  iPlayer = objLineups.Fields("sGuard2").Value				  				  
			  objPlayer.Open "select * from tblPlayers where pid = "&iPlayer&" and NBATEAMID in ("&teamList&") ", objConn,3,3,1
			  if objPlayer.RecordCount > 0 then
				 strSQL = "update tbl_lineups set sGuard2 = 9995, sGuard2Barps = 0, sGuardTip2 = '"&c_Time_1159&"' "&_ 
				          "where ownerID = "&iownerId&" and gameday = #"&currentDate&"#"
				 objConn.Execute strSQL
				 Response.Write "--> Removed sGuard2 for '"&sOwner&"'<br>"
			  end if
			  objPlayer.Close				  
			  
			  objLineups.MoveNext
		   Wend
		   objLineups.Close

		   FuncCall = UpdateGridField(iTeam1, "null")
		   FuncCall = UpdateGridField(iTeam2, "null")
		   
			sAction = ""
		case "Add Teams"
			chkboxCnt   = Request.Form("chkTeamSkedAdd").count
			newTipTime    = Request.Form("newAddGameTime")
			currentDate   = Request.Form("gameDay")				
			splitNBATMID  = Split(Request.Form("chkTeamSkedAdd"),";")
			
			
			if chkboxCnt < 2 then
			   Response.Write "### ERROR ## <br> You must select 2 teams.  You selected "&chkboxCnt&".<br>"
			elseif chkboxCnt > 2 then
			   Response.Write "### ERROR ## <br> You can only select 2 teams.  You selected "&chkboxCnt&".<br>"
			elseif newTipTime = "" then
			   Response.Write "### ERROR ## <br>A new time was not selected. <br>"
			else
			   iTeam1 = splitNBATMID(0)
			   iTeam2 = mid(splitNBATMID(1),3)
			   			   		
               'Response.Write "iTeam1 ="&iTeam1&" <br>"
			   'Response.Write "iTeam2 ="&iTeam2&" <br>"
			   'Response.Write "### SUCCESS ## <br>Add Code. <br>"
			   objWork.Open "select * from tblNBATeams where NBATID = "&iTeam1, objConn,3,3,1
			   sShort1 = objWork.Fields("teamShortName").Value
			   sLong1 = objWork.Fields("teamName").Value
			   objWork.Close
			   
			   objWork.Open "select * from tblNBATeams where NBATID = "&iTeam2, objConn,3,3,1
			   sShort2 = objWork.Fields("teamShortName").Value
			   sLong2 = objWork.Fields("teamName").Value
			   objWork.Close
			   
			   Response.Write "Game Day = "&currentDate&"<br>"
			   Response.Write "Time = "&newTipTime&"<br>"
			   Response.Write "Adding to NBAINDTMSKed  -->  "&sLong1&", "&sLong2&"<br>"
			   
			   strSQL = "insert into NBAINDTMSked (NBATeam, GameDay, GameTime, OppLongName, Opponent) " & _
		                "values ("&iTeam1&",#"&currentDate&"#,'"&newTipTime&"','vs "&sLong2&"','vs "&sShort2&"')"
						
			   objConn.Execute strSQL			   
			   'Response.Write "strSQL = "&strSQL&"<br>"
			   			   			   
			   strSQL = "insert into NBAINDTMSked (NBATeam, GameDay, GameTime, OppLongName, Opponent) " & _
		                "values ("&iTeam2&",#"&currentDate&"#,'"&newTipTime&"','at "&sLong1&"','at "&sShort1&"')"
						
			   objConn.Execute strSQL
			   'Response.Write "strSQL = "&strSQL&"<br>"
			   
			   FuncCall = UpdateGridField(iTeam1, 1)
			   FuncCall = UpdateGridField(iTeam2, 1)
			   
				sAction = ""
			end if
		
		end select	
		
   Function UpdateGridField (pTeam, pgridValue)	
    if pTeam = 1 then
	   strSQL = "update tblGameGrid set ATL = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 2 then
	   strSQL = "update tblGameGrid set BOX = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 3 then
	   strSQL = "update tblGameGrid set BKN = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 4 then
	   strSQL = "update tblGameGrid set CHA = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 5 then
	   strSQL = "update tblGameGrid set CHI = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 6 then
	   strSQL = "update tblGameGrid set CLE = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 7 then
	   strSQL = "update tblGameGrid set DAL = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 8 then
	   strSQL = "update tblGameGrid set DEN = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 9 then
	   strSQL = "update tblGameGrid set DET = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 10 then
	   strSQL = "update tblGameGrid set GSW = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 11 then
	   strSQL = "update tblGameGrid set HOU = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 12 then
	   strSQL = "update tblGameGrid set IND = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 13 then
	   strSQL = "update tblGameGrid set LAC = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 14 then
	   strSQL = "update tblGameGrid set LAL = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 15 then
       strSQL = "update tblGameGrid set MEM = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 16 then
	   strSQL = "update tblGameGrid set MIA = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 17 then
	   strSQL = "update tblGameGrid set MIL = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 18 then
	   strSQL = "update tblGameGrid set MIN = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 19 then
	   strSQL = "update tblGameGrid set NOP = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 20 then
	   strSQL = "update tblGameGrid set NYK = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 21 then
	   strSQL = "update tblGameGrid set OKC = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 22 then
	   strSQL = "update tblGameGrid set ORL = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 23 then
	   strSQL = "update tblGameGrid set PHI = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 24 then
	   strSQL = "update tblGameGrid set PHX = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 25 then
	   strSQL = "update tblGameGrid set POR = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 26 then
	   strSQL = "update tblGameGrid set SAC = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 27 then
	   strSQL = "update tblGameGrid set SAS = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 28 then
	   strSQL = "update tblGameGrid set TOR = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 29 then
	   strSQL = "update tblGameGrid set UTA = "&pgridValue&" where GameDate = #"&currentDate&"#"
	elseif pTeam = 30 then
	   strSQL = "update tblGameGrid set WAS = "&pgridValue&" where GameDate = #"&currentDate&"#"
	else
	   Response.Write "******** ERROR, ERROR, ERROR, ERROR, ERROR..  pTeam value is invalid.  Value = "&pTeam
	end if
	
	objConn.Execute strSQL
	Response.Write "strSQL = "&strSQL&"<br>"
	
   End Function
%>
<!--#include virtual="Common/functions.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2017-2018</title>
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style type="text/css">
.panel-override {
  background-color:white;
  border-color:#354478;
	color:black
}
</style>
<script>
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip();   
});
</script>
</head>
<body>
<script type="text/javascript">
function processTeams(theForm) {
	
	var teamCnt = 2;
	var teamsSelected = 0;
	var timeEntered 
	timeEntered = theForm.newTime.value
	
	for(var i=0; i < theForm.chkTeamSked.length; i++){
		if(theForm.chkTeamSked[i].checked) {
		teamsSelected +=1;
		}
	}

	if(teamsSelected < 2  ) {
		alert("Two Teams Required for Update!" ); 
		return (false);
	}else if(teamsSelected > 2  ) {
		alert("Only Two Teams Can be Selected for Update!" ); 
		return (false); 
	}

	if(!timeEntered) {
		alert("New Time Required for Update!" ); 
		return (false);
	}

	return (true);
}	

function processNewTeams(theForm) {
	
	var teamCnt = 2;
	var teamsSelected = 0;
	var timeEntered 
	timeEntered = theForm.newAddGameTime.value
	
	for(var i=0; i < theForm.chkTeamSkedAdd.length; i++){
		if(theForm.chkTeamSkedAdd[i].checked) {
		teamsSelected +=1;
		}
	}

	if(teamsSelected < 2  ) {
		alert("Two Teams Required for Update!" ); 
		return (false);
	}else if(teamsSelected > 2  ) {
		alert("Only Two Teams Can be Selected for Update!" ); 
		return (false); 
	}

	if(!timeEntered) {
		alert("New Time Required for Update!" ); 
		return (false);
	}

	return (true);
}	

</script>
<!--#include virtual="Common/headerMain.inc"-->
<%
if sAction = "" then
   objRSgames.Open  "select * from qryGameDeadLines", objConn,1,1
%>
<form action="maintainTeamSkeds.asp" name="maintainTeamSkeds" method="POST">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_pid" value="<%=pid%>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Game Date Maintenance</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
				<div class="panel-heading clearfix">
					<h4 class="panel-title">Maintain Tip Times</h4>
				</div>
				<div class="panel-body">
					<table class="table table-striped table-responsive table-bordered table-condensed">
						<tr>
							<td>        
							<select class="form-control input-sm" name="gameDays">
								<% While not objRSgames.EOF %>
								<option value="<%=objRSgames("gameDay")%>"><%=objRSgames.Fields("gameDay")%> </option>
								<% objRSgames.MoveNext
								Wend 
								%>
								</select>
							</td>
						</tr>
					</table>
				</div>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
			<button type="submit" name="Action"  value="Retrieve Teams" class="btn btn-default btn-block  btn-sm"><span class="glyphicon glyphicon-save"></span>&nbsp;Retrieve Teams</button>
		</div>
	</div>
</div>
</form>
<%end if
if sAction = "Retrieve Teams" then 
	 objRSTMSkeds.Open  "SELECT * FROM tblNBATeams INNER JOIN NBAINDTMSKed ON tblNBATeams.NBATID = NBAINDTMSKed.NBATeam " &_ 
	                    "WHERE NBAINDTMSKed.[gameDay]= CDATE('"&currentDate&"') order by GameTime, TeamName ",objConn,3,3,1
%>
<form action="maintainTeamSkeds.asp" name="maintainTeamSkeds" method="POST" onSubmit="return processTeams(this)">
<input type="hidden" name="gameDay" value="<%= objRSTMSkeds.Fields("gameday").Value %>" />
<!--#include virtual="Common/headermain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
				<div class="panel-body">
				<table class="table table-custom table-responsive table-striped table-bordered table-condensed">
						<tr style="background-color:white;">
							<td colspan="3" style="vertical-align:middle;font-weight:bold;text-align:left;background-color:#354478;color:white;">Date: <%= (FormatDateTime(objRSTMSkeds.Fields("gameday").Value,1)) %><small class="badgeEven pull-right">Teams: <%= objRSTMSkeds.RecordCount %></small></td>
						</tr>	
						<tr>
							<td colspan="2" style="text-align:right;width:70%;" colspan="2">ENTER NEW NBA TIP TIME</td>
							<td ><input type="time" class="form-control input-xs" name="newTime"  id="newTime"></td>
						<tr>
					<tr style="color:#354478;">
						<th style="text-align:middle;width:10%;text-align:center;"></th>	
						<th style="text-align:center;width:60%;">Team</th>
						<th style="text-align:middle;width:30%;">Tip Time</th>
					</tr>
						<%
				    While Not objRSTMSkeds.EOF
						%>
						<tr style="background-color:white;">
							<td style="vertical-align:middle;text-align:center;"><input type="checkbox" name="chkTeamSked" value="<%=objRSTMSkeds.Fields("NBATeam").Value%>;"></td>
							<td style="vertical-align:middle;text-align:left;"><%=objRSTMSkeds.Fields("TeamName").Value%>&nbsp;<span><small style="color:red;font-size:10px;font-weight:bold;"><%=objRSTMSkeds.Fields("opponent").Value%></small></span></td>
							<td style="vertical-align:middle;"><%=objRSTMSkeds.Fields("GameTime").Value%></td>
						</tr>	
						<%
							objRSTMSkeds.MoveNext
							Wend
						%>
				</table>
				</div>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-xs-6 col-md-6">
			<button type="submit" value="Update Times" name="Action" class="btn btn-default btn-block btn-sm"><span class="glyphicon glyphicon-save"></span>&nbsp;Update Game Tip TImes</button>
		</div>
		<div class="col-xs-6 col-md-6">
			<button type="submit" value="Delete Teams" name="Action" class="btn btn-danger btn-block btn-sm"><span class="glyphicon glyphicon-save"></span>&nbsp;Delete Teams </button>
		</div>
	</div>
</div>
</form>
<br>
<!--SECTION TO ADD NEW TEAM ENTRIES -->
<%
		objRSnoTMSkeds.Open  "SELECT * FROM tblNBATeams " &_ 
                         "WHERE NOT EXISTS (SELECT 1 from NBAINDTMSKed WHERE NBAINDTMSKed.[gameDay]= CDATE('"&currentDate&"') " &_ 
												 "AND NBAINDTMSKed.NBATeam = tblNBATeams.NBATID) AND tblNBATeams.NBATID > 0",objConn,3,3,1										
		objRSTMSkeds.Close
	  objRSTMSkeds.Open  "SELECT * FROM tblNBATeams INNER JOIN NBAINDTMSKed ON tblNBATeams.NBATID = NBAINDTMSKed.NBATeam " &_ 
	                     "WHERE NBAINDTMSKed.[gameDay]= CDATE('"&currentDate&"') order by GameTime, TeamName ",objConn,3,3,1												 
%>
<form action="maintainTeamSkeds.asp" method="POST" name="frmMain2"  onSubmit="return processNewTeams(this)">
<input type="hidden" name="gameDay" value="<%= objRSTMSkeds.Fields("gameday").Value %>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
				<div class="panel-body">
				<table class="table table-custom table-responsive table-striped table-bordered table-condensed">
						<tr style="background-color:white;">
							<td colspan="2" style="vertical-align:middle;font-weight:bold;text-align:left;background-color:#354478;color:white;">Date: <%= (FormatDateTime(objRSTMSkeds.Fields("gameday").Value,1)) %><small class="badgeEven pull-right">Teams: <%= objRSnoTMSkeds.RecordCount %></small></td>
						</tr>	
						<tr>
							<td style="text-align:middle;width:35%;text-align:center;">ENTER TIP TIME: </td>	
							<td style="text-align:center;width:65%;"><input type="time" class="form-control input-xs" name="newAddGameTime"  id="newAddGameTime"></td>
						<tr>
					<tr style="color:#354478;">
						<th style="text-align:middle;width:35%;text-align:center;"></th>	
						<th style="text-align:center;width:65%;">Team</th>
					</tr>
						<%
							While Not objRSnoTMSkeds.EOF
						%>
					<tr style="background-color:white;">
						<td style="vertical-align:middle;text-align:center;"><input type="checkbox" name="chkTeamSkedAdd" value="<%=objRSnoTMSkeds.Fields("NBATID").Value%>;"></td>
						<td style="vertical-align:middle;text-align:left;"><%=objRSnoTMSkeds.Fields("TeamName").Value%></td>
					</tr>	
						<%
							objRSnoTMSkeds.MoveNext
							Wend
						%>
				</table>
				</div>
			</div>
		</div>
	</div>
</div>
<!--END SECTION TO ADD NEW TEAM ENTRIES -->
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12">
			<button type="submit" value="Add Teams" name="Action" class="btn btn-default btn-block btn-sm"><span class="glyphicon glyphicon-save"></span>&nbsp;Add Teams</button>
		</div>
	</div>
</div>
</form>
<br>
<%end if%>
<%
	objRSgames.close
	objRSTMSkeds.Close
	objRSnoTMSkeds.Close
	Set objRSTMSkeds       = Nothing
	objRSPlayers.Close
	Set objRSPlayers= Nothing
	objConn.Close
	Set ObjConn     = Nothing

%>
</body>
</html>