<!-- #include file="adovbs.inc" -->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear

	strErrorUrl = ""

	Dim objConnSlave, objConnMaster, w_action, teamNum, objrsTest, objStgLast5, objStgBarps, objOpponent, objPosChg, objEmail
	Dim objRSteam, loopcnt
		
	Set objConnStage   = Server.CreateObject("ADODB.Connection")
	Set objConnMaster  = Server.CreateObject("ADODB.Connection")	
	Set objStgLast5    = Server.CreateObject("ADODB.RecordSet")
	Set objStgBarps    = Server.CreateObject("ADODB.RecordSet")
	Set objOpponent    = Server.CreateObject("ADODB.RecordSet")
	Set objPosChg      = Server.CreateObject("ADODB.RecordSet")
	Set objEmail       = Server.CreateObject("ADODB.RecordSet")
	
	
	objConnStage.Open Application("igblDev_ConnectionString")
	objConnStage.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
											"Data Source=igblDev.mdb;" & _
											"Persist Security Info=False"

	objConnMaster.Open Application("lineupstest_ConnectionString")	
	objConnMaster.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
											"Data Source=lineupstest.mdb;" & _
											"Persist Security Info=False;"
	%>
<!--#include virtual="Common/session.inc"-->
<!--#include virtual="Common/headerMain.inc"-->
	<%
			'Response.Write "Value is: " & ownerid & ".<br>"
							
	if Request.Form("action") = "Save Form Data" Then
	
		Response.Write "*** Process Started: "&now()&"<br>"				 
	    'Response.Write "Manage History: "&now()&"<br>"				 
		'strSQL = "delete from zLast5_History where gamedate > date() - 7"
		'objConnStage.Execute strSQL
		
		'strSQL = "insert into zLast5_History select * FROM Monster_Stage_Last5 where gamedate > date() - 7"
		'objConnStage.Execute strSQL
		
		'Response.Write "....Manage History done: "&now()&"<br><br>"
		
		'################################
		'########### tblLast5 ###########
		'################################
		strSQL = "delete from tblLast5"
		objConnMaster.Execute strSQL
				
		'Pull Data From Stage Database
		'objStgLast5.Open "SELECT * FROM Stage_Last5 order by Player, gamedate desc", objConnStage,3,3,1
		objStgLast5.Open "SELECT * FROM Monster_Stage_Last5 order by Player, gamedate desc", objConnStage,3,3,1
		Response.Write "....Record Count From Stage tblLast5 = "&objStgLast5.RecordCount&"<br>"

		lnLoopCt      = 0
		totBarps      = 0 
		divisor       = 0
		l5Count       = 0
		lrowsInserted = 0
		wtot3dbl      = 0
		
        strSQL = "update tblPlayers set l5barps = 0, numTdbls = 0"
		objConnMaster.Execute strSQL
		
		While Not objStgLast5.EOF
			lnLoopCt= lnLoopCt + 1
			wGameDay = objStgLast5.Fields("gameDate").Value
			wPlayer  = replace(replace(objStgLast5.Fields("Player").Value, "'", ""), ".", "")
			wFirst   = Left(wPlayer, instr(wPlayer," ") - 1)
			wLast    = Right(wPlayer, Len(wPlayer) - instr(wPlayer, " "))
			wPos     = objStgLast5.Fields("Pos").Value
			w3       = objStgLast5.Fields("x3P").Value
			wR       = objStgLast5.Fields("TRB").Value
			wA       = objStgLast5.Fields("AST").Value
			wS       = objStgLast5.Fields("STL").Value
			wB       = objStgLast5.Fields("BLK").Value
			wT       = objStgLast5.Fields("TOV").Value
			wP       = objStgLast5.Fields("PTS").Value
			wMP      = objStgLast5.Fields("MP").Value
			wBarps   = objStgLast5.Fields("BARPS").Value  
			wTdbl    = objStgLast5.Fields("TDBL").Value
			
			if ISNULL(objStgLast5.Fields("usage").Value) then
			   wU   = 0
			else
			   wU   = objStgLast5.Fields("usage").Value
			end if
	
			objOpponent.Open 	"SELECT t4.teamShortName " &_
								"from tblPlayers t, tblNBATeams t2, tblleaguesetup t3, tblNBATeams t4 " &_
								"where t.firstname = '"&wFirst&"' and t.lastname = '"&wLast&"' " &_
								"and t.NbaTeamID = t2.NbaTid " &_
								"and t2.TeamName = t3.TeamName " &_
								"and t3.gamedate = #"&wGameDay&"# " &_
								"and t3.Opponent = t4.TeamName ",objConnMaster,3,3,1
							 
			if objOpponent.RecordCount > 0 then
			   wOpp	= objOpponent.Fields("teamShortName").Value
			else
			   wOpp = "###"
			end if
			objOpponent.Close
			
			'First time thru, assign values to wHold
			if lnLoopCt = 1 then			    
				wHoldFirst = wFirst
				wHoldLast  = wLast
				wHoldPos   = wPos
			end if
			
			if wHoldFirst = wFirst AND wHoldLast = wLast then				
				l5Count = l5Count + 1
				
				wtot3dbl = wtot3dbl + wTdbl
				if l5Count <= 5 then
					'totBarps = cint(totBarps) + cint(wBarps)
					totBarps = totBarps + wBarps
				end if
			else	
			    wRetcd = UpdateL5BarpAvg()	
			end if			
						
			if l5Count <= 5 then
			   		
                if l5Count <= 5 then
				   wlast5game = 1
                else
				   wlast5game = 0
                end if			
				
				strSQL = "insert into tblLast5 " & _
		         "(gameDate, first, last, x3P, TRB, AST, STL, BLK, TOV, PTS, MP, BarpTot, Opp, USAGE, TDBL, L5Game) " & _
                  "values (#"&wGameDay&"#,'"&wFirst&"','"&wLast&"',"&w3&","&wR&","&wA&","&wS&","&wB&","&wT&","&wP&"," & _ 
				           wMP&","&wBarps&",'"&wOpp&"',"&wU&","&wTdbl&","&wlast5game&")"
				 
				 objConnMaster.Execute strSQL	
				 lrowsInserted = lrowsInserted + 1
				 
				 'if l5Count = 1 then
			     '  Response.Write "sqlStr.  "&strSQL&"<br>"
			     'end if
		    end if
		    								
			objStgLast5.MoveNext			
		Wend
		objStgLast5.Close
		
		'Handle Last Player
		if lnLoopCt > 0 then
		   wRetcd = UpdateL5BarpAvg()
		end if
		
				
		'strSQL = "UPDATE tblPlayers t SET l5barps = 0 WHERE l5barps <> 0 " & _
        '         "AND not exists (select 1 from tblLast5 t2 where t2.first = t.firstName and t2.last = t.lastName)"
		'objConnMaster.Execute strSQL
		
		Response.Write "....Records written to linesupstest = "&lrowsInserted&"<br>"
		Response.Write "....Last5 with Last 5 Avg Done: "&now()&"<br><br>"
		'#################################
		'########### tbl_barps ###########
		'#################################
		strSQL = "delete from tbl_barps"
		objConnMaster.Execute strSQL
		
		'Pull Data From Stage Database
		objStgBarps.Open "SELECT * FROM Stage_Barps", objConnStage,3,3,1
		Response.Write "....Record Count From Stage tbl_barps = "&objStgBarps.RecordCount&"<br>"

	
		While Not objStgBarps.EOF
			wRank  = objStgBarps.Fields("rank").Value
			wBarps = objStgBarps.Fields("barps").Value
			wPlayer= replace(replace(objStgBarps.Fields("Name").Value, "'", ""), ".", "")
			wFirst = Left(wPlayer, instr(wPlayer," ") - 1)
			wLast  = Right(wPlayer, Len(wPlayer) - instr(wPlayer, " "))
			wTeam  = objStgBarps.Fields("team").Value
			wgp    = objStgBarps.Fields("gp").Value
			wmin   = objStgBarps.Fields("min").Value
			wppg   = objStgBarps.Fields("ppg").Value
				w3   = objStgBarps.Fields("three").Value
			wR     = objStgBarps.Fields("reb").Value
			wA     = objStgBarps.Fields("ast").Value
			wS     = objStgBarps.Fields("stl").Value
			wB     = objStgBarps.Fields("blk").Value
			wT     = objStgBarps.Fields("to").Value
			wU     = objStgBarps.Fields("usage").Value
			
			strSQL = "insert into tbl_barps " & _
		   "(rank, barps, first, last, team, gp, min, ppg, three, reb, ast, stl, blk, to,usage) " & _
		   "values ("&wRank&","&wBarps&",'"&wFirst&"','"&wLast&"','"&wTeam&"',"&wgp&","&wmin&","&wppg& _
		            ","&w3&","&wR&","&wA&","&wS&","&wB&","&wT&","&wU&")"
		    
			objConnMaster.Execute strSQL			
			'Response.Write "strSQL = "&strSQL&"<br>"
			
			objStgBarps.MoveNext			
		Wend
		objStgBarps.Close
		
		'wRetcd = EmailPositionChg()
		Response.Write "....Barps Done: "&now()&"<br>"
		Response.Write "*** Processed Ended: "&now()&"<br>"					 	
	end if
	

	objConnStage.Close
	objConnMaster.Close
	
	Function UpdateL5BarpAvg ()
	
		if l5Count > 5 then
			divisor = 5
		else
			divisor = l5Count
		end if
				
		avgBarps = totBarps/divisor
				
		'if wHoldPos = "C" then
		'   wNewPos = "CEN"
		'elseif wHoldPos = "F/C" OR wHoldPos = "SF/PF/C" then
		'   wNewPos = "F-C"
		''elseif wHoldPos = "F" OR wHoldPos = "SF/PF" then
		'   wNewPos = "FOR"
    '    elseif wHoldPos = "G/F" OR wHoldPos = "PG/SG/SF/PF" OR wHoldPos = "SG/PF" OR wHoldPos = "SG/SF" OR wHoldPos = "SG/SF/PF" then
		'   wNewPos = "G-F"
		'elseif wHoldPos = "G" OR wHoldPos = "PG/SG" OR wHoldPos = "SG" then
		'   wNewPos = "GUA"
		'else
		 wNewPos = "XXX"
		'   Response.Write " ### POSITION NOT ASSIGNED for "&wHoldFirst&" "&wHoldLast&".  ESPN position = "&wHoldPos
		'end if
		
		
		
		
		
		strSQL = "update tblPlayers set l5barps = "&avgBarps&", numTdbls = "&wtot3dbl&", espnPos = '"&wHoldPos&"', NewPos = '"&wNewPos&"' " &_
		         "where firstname = '"&wHoldFirst&"' and lastname = '"&wHoldLast&"'"
		objConnMaster.Execute strSQL
		
		'Response.Write "strSQL = "&strSQL&"<br>"
		'Response.Write "*******************************************<br>"
		
				
		wHoldFirst = wFirst
		wHoldLast  = wLast
		wHoldPos   = wPos
		wtot3dbl   = wTdbl
		l5Count = 1
		totBarps = wBarps

	End Function
	
    Function EmailPositionChg ()
	
	    objPosChg.Open "SELECT * FROM tblPlayers where newpos is not null and newpos <> pos order by lastname", objConnMaster,3,3,1
		if objPosChg.RecordCount > 0 then
		   Response.Write tst_msg
		   sPositionChgString = "The Following players positions have changed: <br><br>"	   
		   While Not objPosChg.EOF
		      wfname = objPosChg.Fields("firstname").Value
			  wlname = objPosChg.Fields("lastname").Value
			  wNewPos = objPosChg.Fields("newpos").Value
			  
		      sPositionChgString = sPositionChgString&wfname&" "&wlname&" new position is "&wNewPos&"<br>"
		      objPosChg.MoveNext
		   Wend
		   
		   objEmail.Open  "Select * from tblOwners where ownerid <> 99 ", objConnMaster
		   email_to = objEmail.Fields("HomeEmail").Value
		   objEmail.MoveNext
		
		   While Not objEmail.EOF				    
		      email_to = email_to & "," & objEmail.Fields("HomeEmail").Value
			  objEmail.MoveNext
		   Wend
		   objEmail.Close
		   				
		   email_subject= "PLAYER POSITION CHANGE"  'You can put whatever subject here
		   host         = "mail.igbl.org"   'The mail server name. (Commonly mail.yourdomain.xyz if your mail is hosted with HostMySite) 
		   username     = "igbl_commish@igbl.org"  'A valid email address you have setup 
		   from_address ="igbl_commish@igbl.org" 'If your mail is hosted with HostMySite this has to match the email address above 
		   password     = "SpursRtheChamps2014" 'Password for the above email address
		   reply_to     = "noReply@igbl.org"  'Enter the email you want customers to reply to
		   port = "25" 'This is the default port. Try port 50 if this port gives you issues and your mail is hosted with HostMySite

		   Set ObjSendMail = CreateObject("CDO.Message")
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = host
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = port
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
		   ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
		   ObjSendMail.Configuration.Fields.Update

		   email_message       = sPositionChgString
		   email_message       = email_message &"<br>http://www.igbl.org"	
			
		   email_to     = "dennis_myers@igbl.org, fred_curry@igbl.org"  'Enter the email you want to send the form to			
		   ObjSendMail.To      = email_to
		   ObjSendMail.Subject = email_subject
		   ObjSendMail.From    = from_address    
		   ObjSendMail.HTMLBody= email_message   
		   ObjSendMail.Send
		   set ObjSendMail     = Nothing
		   		   
		   'Update new positions
		   'strSQL = "update tblplayers set pos = newpos where newpos is not null and newpos <> pos"
		   'objConnMaster.Execute strSQL
		   
		end if
		
		objPosChg.Close
			
	End Function		

%>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="Dee M. Myers" content="">
<title>Generate Results</title>
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
.bs-callout-success {
    border-left-color: #000000;
    padding: 10px;
    border-left-width: 4px;
    border-radius: 3px;
    background-color: white;
}
</style>
</head>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Transfer Last 5 Stats</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
  <div class="bs-callout bs-callout-success">
    <h4>Transfer Game Data </h4>
    <ol>
		<li>Retrieve Game Stats for Date Range BBM</li>
		<li>Export to Excel File</li>
		<li>Remove Own/Owner Columns</li>
		<li>Add Date Column & Enter Today's Game Date</li>
		<li>Download IGBLDev Database </li>
		<li>Add Data to Monster_Stage_Last5 table</li>
		<li>Save/Upload Database to Server</li>
		<li>Click Button to Copy Data to Main Database</li>
	</ol>
  </div>
</div>
<br>
<body bgcolor="#FFFFF7">
<form action="transferLast5.asp" method="POST">
<input type="hidden" name="action" value="Save Form Data">
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<button class="btn btn-lg btn-default btn-block" value="Run BuildBox" name="Submit" type="submit">Transfer Game Stats Data</button>
		</div>
	</div>
</div>
</form>
</body>
</html>