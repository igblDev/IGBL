<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%

On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim objConn,objRSLineups,objRSTeamLogos,ownerid

	ownerid = session("ownerid")	
	
  if ownerid = "" then
    	GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if	
	
	Set objConn             = Server.CreateObject("ADODB.Connection")
	Set objRSLineups        = Server.CreateObject("ADODB.RecordSet")

	Set objNextRun          = Server.CreateObject("ADODB.RecordSet")
	Set objRSLineupsPics    = Server.CreateObject("ADODB.RecordSet")
	Set objRSTeamLogos      = Server.CreateObject("ADODB.RecordSet")
	Set objRSBarps          = Server.CreateObject("ADODB.RecordSet")
	
	objConn.Open Application("lineupstest_ConnectionString")


	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
				  
				  

%>
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2016-2017</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
black{
	color:black;
}
orange{
	color:darkorange;
}
red {
	color: red;
}
yellow {
	color:yellow;
}
green {
	color:#468847;
}
td {
    vertical-align: middle;		
		font-size:12px;
		color:black;
}
.panel-heading{

}
small {
font-size:10px;
color:gray;
}
.alert-success {
    border-color: #468847;;
}
/* tab color */

.nav-tabs>li>a {
	background-color: darkorange;
	border-color: darkorange;
	/*50%*/
	color: #FFFFFF;
	font-weight: bold;
	padding-top: 3px;
	padding-bottom: 1px;
}

/* active tab color */

.nav-tabs>li.active>a,
.nav-tabs>li.active>a:hover,
.nav-tabs>li.active>a:focus {
	color: #000;
	background-color: #fff;
	border: 1px solid #354478;
}

/* hover tab color */

.nav-tabs>li>a:hover,
.nav-tabs>li>a:focus {
	border-color: #354478;
	background-color: #C5DDF1;
	color: #000;
}

/* left tabs */

.nav-tabs,
.nav-pills {
	text-align: left;
}

/* line below tabs */

.nav-tabs {
	border-bottom: 2px solid #354478;
}
.panel-body {
    padding: 10px;
    padding-top: 15px;
    padding-right: 10px;
    padding-bottom: 10px;
    padding-left: 10px;
}
.panel-matchups {
  background-color:#d9edf7;
  border-color:#354478;
	color:#354478;
	border-radius:0px;
}
.alert-info {
    color: #31708f;
    background-color: #d9edf7;
    border-color: #d9edf7;
}
.badge2 {
    display: inline-block;
    min-width: 10px;
    padding: 3px 7px;
    font-size: 10px;
    font-weight: 700;
    line-height: 1;
    color: #fff;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color: #354478;
    border-radius: 10px;
}
.badge3 {
    display: inline-block;
    min-width: 10px;
    padding: 3px 7px;
    font-size: 10px;
    font-weight: 700;
    line-height: 1;
    color:white;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color:#354478;
    border-radius: 10px;
}
.badgeUp {
    display:inline-block;
    min-width:10px;
    padding:3px 7px;
    font-size:10px;
    font-weight:700;
    line-height:1;
    color: white;
    text-align:center;
    white-space:nowrap;
    vertical-align:baseline;
    background-color:#468847;
    border-radius:10px;
}
.badgeDown {
    display:inline-block;
    min-width:10px;
    padding:3px 7px;
    font-size:10px;
    font-weight:700;
    line-height:1;
    color: white;
    text-align:center;
    white-space:nowrap;
    vertical-align:baseline;
    background-color:red;
    border-radius:10px;
}
.badgeEven {
    display:inline-block;
    min-width:10px;
    padding:3px 7px;
    font-size:10px;
    font-weight:700;
    line-height:1;
    color: white;
    text-align:center;
    white-space:nowrap;
    vertical-align:baseline;
    background-color:gold;
    border-radius:10px;
}
.panel-override {
  color:#354478;
  background-color:white;
  border-color:#354478;
	border-radius:0px;
}
small {
	color:black;
}
</style>
</head>
<body>
<% 

	Set objRSMatchups= Server.CreateObject("ADODB.RecordSet")
	Set objRSBarps   = Server.CreateObject("ADODB.RecordSet")
	Set objRSPics    = Server.CreateObject("ADODB.RecordSet")   
	Set objRSTeamLogos   = Server.CreateObject("ADODB.RecordSet")
	Set objRSNBASked  = Server.CreateObject("ADODB.RecordSet")
	
	objRSMatchups.Open "SELECT * FROM qry_matchupLineups_stagger", objConn,3,3,1
	gameday = objRSMatchups.Fields("gameday").Value
	dim loopcnt 
	loopcnt = 1
	

	%>
<form>
<div class="container">
	<div class="row">		
		<div class="col-md-12 col-sm-12 col-xs-12">			
			<span style=""><strong>Game Day Matchups</strong></span><br>
			<%=(FormatDateTime(gameday,1))%>
		</div>
	</div>
</div>
</br>
<!--#include virtual="Common/headerMain.inc"-->

				<% if objRSMatchups.RecordCount > 0 then %>
				<% 
				
				While Not objRSMatchups.EOF
	
				hometeam      = objRSMatchups.Fields("HomeTeam").Value
				hometeamshort = objRSMatchups.Fields("HomeTeamShort").Value
				homeowner     = objRSMatchups.Fields("HomeOwner").Value
				hometeampen   = objRSMatchups.Fields("HomeTeamPen").Value
				awayteam      = objRSMatchups.Fields("VisitingTeam").Value
				awayteamshort = objRSMatchups.Fields("VisitingTeamShort").Value
				awayowner     = objRSMatchups.Fields("VisitingOwner").Value
				awayteampen   = objRSMatchups.Fields("VisitingTeamPen").Value
				
				objRSTeamLogos.Open "Select TeamLogo  from TBLOWNERS where  " & homeowner  & " = ownerID " , objConn,3,3,1
				homeLogo = objRSTeamLogos.Fields("TeamLogo").Value
				objRSTeamLogos.Close		
				
				objRSTeamLogos.Open "Select TeamLogo  from TBLOWNERS where  " & awayowner  & " = ownerID " , objConn,3,3,1
				awayLogo = objRSTeamLogos.Fields("TeamLogo").Value
				objRSTeamLogos.Close		
				
				homerec  = objRSMatchups.Fields("HomeRecord").Value
				awayrec  = objRSMatchups.Fields("VisRecord").Value
				homeStamp= objRSMatchups.Fields("HomeStamp").Value
				awayStamp= objRSMatchups.Fields("Timestamp").Value
				'******************************
				'** Retrieving Starters PID  **
				'******************************
				sCenterAway   = objRSMatchups.Fields("sCenter").Value
				sForward1Away = objRSMatchups.Fields("sForward").Value
				sForward2Away = objRSMatchups.Fields("sForward2").Value
				sGuard1Away   = objRSMatchups.Fields("sGuard").Value
				sGuard2Away   = objRSMatchups.Fields("sGuard2").Value	
				sCenterHome   = objRSMatchups.Fields("HomeC").Value
				sForward1Home = objRSMatchups.Fields("HomeF").Value
				sForward2Home = objRSMatchups.Fields("HomeF2").Value
				sGuard1Home   = objRSMatchups.Fields("HomeG").Value
				sGuard2Home   = objRSMatchups.Fields("HomeG2").Value
				lineupTimeChk = time() - 1/24	
				
				cTip          = objRSMatchups.Fields("sCenterTip").Value
				if len(objRSMatchups.Fields("sCenterTip").Value) = 10 then
					cTip = Left(objRSMatchups.Fields("sCenterTip").Value,4) & Right(objRSMatchups.Fields("sCenterTip").Value,3)
				else
					cTip = Left(objRSMatchups.Fields("sCenterTip").Value,5) & Right(objRSMatchups.Fields("sCenterTip").Value,3)
				end if	
				
				
				for1Tip       = objRSMatchups.Fields("sForwardTip").Value
				if len(objRSMatchups.Fields("sForwardTip").Value) = 10 then
					for1Tip = Left(objRSMatchups.Fields("sForwardTip").Value,4) & Right(objRSMatchups.Fields("sForwardTip").Value,3)
				else
					for1Tip = Left(objRSMatchups.Fields("sForwardTip").Value,5) & Right(objRSMatchups.Fields("sForwardTip").Value,3)
				end if	

				for2Tip       = objRSMatchups.Fields("sForwardTip2").Value
				if len(objRSMatchups.Fields("sForwardTip2").Value) = 10 then
					for2Tip = Left(objRSMatchups.Fields("sForwardTip2").Value,4) & Right(objRSMatchups.Fields("sForwardTip2").Value,3)
				else
					for2Tip = Left(objRSMatchups.Fields("sForwardTip2").Value,5) & Right(objRSMatchups.Fields("sForwardTip2").Value,3)
				end if	

				gua1Tip       = objRSMatchups.Fields("sGuardTip").Value
				if len(objRSMatchups.Fields("sGuardTip").Value) = 10 then
					gua1Tip = Left(objRSMatchups.Fields("sGuardTip").Value,4) & Right(objRSMatchups.Fields("sGuardTip").Value,3)
				else
					gua1Tip = Left(objRSMatchups.Fields("sGuardTip").Value,5) & Right(objRSMatchups.Fields("sGuardTip").Value,3)
				end if	

				gua2Tip       = objRSMatchups.Fields("sGuardTip2").Value
				if len(objRSMatchups.Fields("sGuardTip2").Value) = 10 then
					gua2Tip = Left(objRSMatchups.Fields("sGuardTip2").Value,4) & Right(objRSMatchups.Fields("sGuardTip2").Value,3)
				else
					gua2Tip = Left(objRSMatchups.Fields("sGuardTip2").Value,5) & Right(objRSMatchups.Fields("sGuardTip2").Value,3)
				end if	

				hCenTip       = objRSMatchups.Fields("HomeCTip").Value
				if len(objRSMatchups.Fields("HomeCTip").Value) = 10 then
					hCenTip = Left(objRSMatchups.Fields("HomeCTip").Value,4) & Right(objRSMatchups.Fields("HomeCTip").Value,3)
				else
					hCenTip = Left(objRSMatchups.Fields("HomeCTip").Value,5) & Right(objRSMatchups.Fields("HomeCTip").Value,3)
				end if	

				hForTip       = objRSMatchups.Fields("HomeFTip").Value
				if len(objRSMatchups.Fields("HomeFTip").Value) = 10 then
					hForTip = Left(objRSMatchups.Fields("HomeFTip").Value,4) & Right(objRSMatchups.Fields("HomeFTip").Value,3)
				else
					hForTip = Left(objRSMatchups.Fields("HomeFTip").Value,5) & Right(objRSMatchups.Fields("HomeFTip").Value,3)
				end if	

				hFor2Tip      = objRSMatchups.Fields("HomeF2Tip").Value
				if len(objRSMatchups.Fields("HomeF2Tip").Value) = 10 then
					hFor2Tip = Left(objRSMatchups.Fields("HomeF2Tip").Value,4) & Right(objRSMatchups.Fields("HomeF2Tip").Value,3)
				else
					hFor2Tip = Left(objRSMatchups.Fields("HomeF2Tip").Value,5) & Right(objRSMatchups.Fields("HomeF2Tip").Value,3)
				end if	

				hGuaTip       = objRSMatchups.Fields("HomeGTip").Value
				if len(objRSMatchups.Fields("HomeGTip").Value) = 10 then
					hGuaTip = Left(objRSMatchups.Fields("HomeGTip").Value,4) & Right(objRSMatchups.Fields("HomeGTip").Value,3)
				else
					hGuaTip = Left(objRSMatchups.Fields("HomeGTip").Value,5) & Right(objRSMatchups.Fields("HomeGTip").Value,3)
				end if	

				hGua2Tip      = objRSMatchups.Fields("HomeG2Tip").Value 
				if len(objRSMatchups.Fields("HomeG2Tip").Value) = 10 then
					hGua2Tip = Left(objRSMatchups.Fields("HomeG2Tip").Value,4) & Right(objRSMatchups.Fields("HomeG2Tip").Value,3)
				else
					hGua2Tip = Left(objRSMatchups.Fields("HomeG2Tip").Value,5) & Right(objRSMatchups.Fields("HomeG2Tip").Value,3)
				end if	

				ACenBarps  		= 0
				AFor1Barps 		= 0
				AFor2Barps 		= 0
				AGua1Barps 		= 0
				AGua2Barps 		= 0
				HCenBarps  		= 0
				HFor1Barps 		= 0
				HFor2Barps 		= 0
				HGua1Barps 		= 0
				HGua2Barps 		= 0
				
				objRSPics.Open "Select image,firstName,lastName,NBATeamID from tblPlayers where  " & sCenterAway  & " = PID " , objConn,3,3,1
				CenterAwayNBATM= objRSPics.Fields("NBATeamID").Value
				ACenPic       = objRSPics.Fields("image").Value
				CenterAway    = objRSPics.Fields("lastName").Value
				CenterAwayNF  = left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				CenterAwayOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close  
				
				'Response.Write "Team ID = "&CenterAwayNBATM&"<br>"
				'Response.Write "Opponent Tip = "&CenterAwayOpp&"<br>"
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sCenterAway  & " = PID " , objConn,3,3,1
				ACenBarps     = objRSBarps.Fields("barps").Value
				ACenTM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select image,firstName,lastName,NBATeamID   from tblplayers where " & sForward1Away & " = PID" , objConn,3,3,1
				Forward1AwayNBATM= objRSPics.Fields("NBATeamID").Value
				AFor1Pic      = objRSPics.Fields("image").Value
				Forward1Away  = objRSPics.Fields("lastName").Value
				Forward1AwayNF= left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Forward1AwayOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close				

				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward1Away & " = PID " , objConn,3,3,1
				AFor1Barps    = objRSBarps.Fields("barps").Value
				AFor1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select image,firstName,lastName,NBATeamID   from tblplayers where " & sForward2Away &" = PID" , objConn,3,3,1
				Forward2AwayNBATM= objRSPics.Fields("NBATeamID").Value
				AFor2Pic      = objRSPics.Fields("image").Value
				Forward2Away  = objRSPics.Fields("lastName").Value
				Forward2AwayNF= left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Forward2AwayOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close				
				objRSPics.Close 

				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward2Away & " = PID " , objConn,3,3,1
				AFor2Barps    = objRSBarps.Fields("barps").Value
				AFor2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select image,firstName,lastName,NBATeamID   from tblplayers where  " & sGuard1Away & " = PID" , objConn,3,3,1
				Guard1AwayNBATM= objRSPics.Fields("NBATeamID").Value
				AGua1Pic      = objRSPics.Fields("image").Value
				Guard1Away    = objRSPics.Fields("lastName").Value
				Guard1AwayNF  = left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Guard1AwayOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close 
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard1Away & " = PID " , objConn,3,3,1
				AGua1Barps    = objRSBarps.Fields("barps").Value
				AGua1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select image,firstName,lastName,NBATeamID   from tblplayers where " & sGuard2Away & "  = PID" , objConn,3,3,1
				Guard2AwayNBATM= objRSPics.Fields("NBATeamID").Value
				AGua2Pic      = objRSPics.Fields("image").Value
				Guard2Away    = objRSPics.Fields("lastName").Value
				Guard2AwayNF  = left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Guard2AwayOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard2Away & " = PID " , objConn,3,3,1
				AGua2Barps    = objRSBarps.Fields("barps").Value
				AGua2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
		
				'******************************************************************************************************
				'**** RETRIEVING THE BARPS AD THE IMAGES FOR THE STARTING PLAYERS ON THE HOME TEAM
				'******************************************************************************************************
				
				objRSPics.Open "Select image,firstName,lastName,NBATeamID from tblplayers where " & sCenterHome & "  = PID" , objConn,3,3,1
				CenterHomeNBATM= objRSPics.Fields("NBATeamID").Value
				HCenPic     = objRSPics.Fields("image").Value
				CenterHomeN = objRSPics.Fields("lastName").Value
				CenterHomeNF= left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				CenterHomeOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sCenterHome & " = PID " , objConn,3,3,1
				HCenBarps  = objRSBarps.Fields("barps").Value
				HCenTM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select image,firstName,lastName,NBATeamID from tblplayers where " & sForward1Home & " = PID" , objConn,3,3,1
				Forward1HomeNBATM= objRSPics.Fields("NBATeamID").Value
				HFor1Pic = objRSPics.Fields("image").Value
				Forward1HomeN = objRSPics.Fields("lastName").Value
				Forward1HomeNF = left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Forward1HomeOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward1Home & " = PID " , objConn,3,3,1
				HGua2HomeNBATM= objRSPics.Fields("NBATeamID").Value
				HFor1Barps  = objRSBarps.Fields("barps").Value
				HFor1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select image,firstName,lastName,NBATeamID from tblplayers where " & sForward2Home & " = PID" , objConn,3,3,1
				Forward2HomeBATM= objRSPics.Fields("NBATeamID").Value
				HFor2Pic = objRSPics.Fields("image").Value
				Forward2HomeN = objRSPics.Fields("lastName").Value
				Forward2HomeNF= left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Forward2HomeOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close	
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sForward2Home & " = PID " , objConn,3,3,1
				HFor2Barps = objRSBarps.Fields("barps").Value
				HFor2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
				objRSPics.Open "Select image,firstName,lastName,NBATeamID from tblplayers where " & sGuard1Home & " = PID" , objConn,3,3,1
				Guard1HomeNBATM= objRSPics.Fields("NBATeamID").Value
				HGua1Pic = objRSPics.Fields("image").Value
				Guard1HomeN = objRSPics.Fields("lastName").Value
				Guard1HomeNF= left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Guard1HomeOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close	
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard1Home & " = PID " , objConn,3,3,1
				HGua2HomeNBATM= objRSPics.Fields("NBATeamID").Value
				HGua1Barps = objRSBarps.Fields("barps").Value
				HGua1TM = objRSBarps.Fields("team").Value
				objRSBarps.Close

				objRSPics.Open "Select image,firstName,lastName,NBATeamID from tblplayers where " & sGuard2Home & " = PID" , objConn,3,3,1
				Guard2HomeNBATM= objRSPics.Fields("NBATeamID").Value
				HGua2Pic      = objRSPics.Fields("image").Value
				Guard2HomeN   = objRSPics.Fields("lastName").Value
				Guard2HomeNF  = left(objRSPics.Fields("firstName").Value,1)
				objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSPics.Fields("NBATeamID").Value, objConn,3,3,1
				Guard2HomeOpp = objRSNBASked.Fields("opponent").value
				objRSNBASked.Close
				objRSPics.Close
				
				objRSBarps.Open "SELECT barps,team FROM qry_tblbarps where  " & sGuard2Home & " = PID " , objConn,3,3,1
				HGua2Barps = objRSBarps.Fields("barps").Value
				HGua2TM = objRSBarps.Fields("team").Value
				objRSBarps.Close
				
	
				htBarps = cDbl(HGua1Barps)  + cDbl(HGua2Barps) + cDbl(HFor1Barps)  + cDbl(HFor2Barps) + cDbl(HCenBarps)
				atBarps = cDbl(AGua1Barps)  + cDbl(AGua2Barps) + cDbl(AFor1Barps)  + cDbl(AFor2Barps) + cDbl(ACenBarps)		

				%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm12 col-xs-12">
				<div class="panel panel-override">
        <table class="table table-condensed table-custom table-responsive table-bordered">
          <tr  align="center">
            <% if awayteampen = true Then %>
            	<th  style="text-align:center;"><%=awayteam%><br><green><%=awayrec%></green></th>
            <% else %> 
            	<th style="width:50%;text-align:center;"><%=awayteam%><br><green><%=awayrec%></green></th>
            <%end if %>
            <TD style="background-color:#f3f3f3;color:354478;vertical-align:middle;text-align:center" rowspan="2"><strong>VS</strong></TD>
						<% if hometeampen = true Then %>
							<th style="text-align:center;"><%=hometeam%><br><green><%=homerec%></green></th>          
						<% else %> 
							<th  style="width:50%;text-align:center;"><%=hometeam%><br><green><%=homerec%></green></th>
            <%end if %>
					</tr>
          <tr  style="background-color:white" align="center">
						<% if awayteampen = true Then %>
							<td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="images/penalty.png"></td>
						<% else %>
	            <td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="<%=awayLogo%>"></td>
						<%end if %>
						<% if hometeampen = true Then %>
						  <td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="images/penalty.png"></td>
						<% else %>
						  <td style="width:50%" align="center"><img style="max-height:100px; max-width:100px" class="img-responsive" border="0" src="<%=homeLogo%>"></td>
						<%end if %>
          </tr>
					<!--<tr >
            <th style="text-align:center;color:black;" colspan="4" colspan="2">Score Projections</th>
          </tr>
          <tr>
						<%if cint(atBarps) > cint(htBarps) then%>
						<td align="center"><span class="badgeUp"><%=atBarps%></span></td>
						<%elseif cint(htBarps) > cint(atBarps) then%>
						<td align="center"><span class="badgeDown"><%=atBarps%></span></td>
						<%else%>
						<td align="center"><span class="badgeEven"><%=atBarps%></span></td>
						<%end if%>
						<td style="text-align:center;background-color:black;"><strong></strong></td>
						<%if cint(atBarps) > cint(htBarps) then%>
						<td align="center"><span class="badgeDown"><%=htBarps%></span></td>
						<%elseif cint(htBarps) > cint(atBarps) then%>
						<td align="center"><span class="badgeUp"><%=htBarps%></span></td>
						<%else%>
						<td align="center"><span class="badgeEven"><%=htBarps%></span></td>
						<%end if%>					
          </tr>-->
					<tr style="vertical-align:middle;text-align:center">
						<td><a href="playerprofile.asp?pid=<%=sCenterAway%>"><%=CenterAwayNF%>.&nbsp;<%=CenterAway %></a></td>
						<td style="background-color:#f3f3f3;color:#354478"><strong>C</strong></td>
            <td><a href="playerprofile.asp?pid=<%=sCenterHome%>"><%=CenterHomeNF%>.&nbsp;<%=CenterHomeN %></a></td>
					</tr>
					<tr style="vertical-align:middle;text-align:center">
						<td><a href="playerprofile.asp?pid=<%=sForward1Away%>"><%=Forward1AwayNF %>.&nbsp;<%=Forward1Away %></a></td>
						<td style="background-color:#f3f3f3;color:#354478"><strong>F</strong></td>
            <td><a href="playerprofile.asp?pid=<%=sForward1Home%>"><%=Forward1HomeNF%>.&nbsp;<%=Forward1HomeN %></a></td>
					</tr>
					<tr style="vertical-align:middle;text-align:center">
            <td><a href="playerprofile.asp?pid=<%=sForward2Away%>"><%=Forward2AwayNF %>.&nbsp;<%=Forward2Away %></td>
						<td style="background-color:#f3f3f3;color:#354478"><strong>F</strong></td>
            <td><a href="playerprofile.asp?pid=<%=sForward2Home%>"><%=Forward2HomeNF%>.&nbsp;<%=Forward2HomeN %></td>
					</tr>
					<tr style="vertical-align:middle;text-align:center">
            <td><a href="playerprofile.asp?pid=<%=sGuard1Away%>"><%=Guard1AwayNF%>.&nbsp;<%=Guard1Away %></a></td>
						<td style="background-color:#f3f3f3;color:#354478"><strong>G</strong></td>
            <td><a href="playerprofile.asp?pid=<%=sGuard1Home%>"><%=Guard1HomeNF%>.&nbsp;<%=Guard1HomeN %></a></td>
					</tr>
					<tr style="vertical-align:middle;text-align:center">
            <td><a href="playerprofile.asp?pid=<%=sGuard2Away%>"><%=Guard2AwayNF%>.&nbsp;<%=Guard2Away %></a></td>
						<td style="background-color:#f3f3f3;color:#354478"><strong>G</strong></td>
            <td><a href="playerprofile.asp?pid=<%=sGuard2Home%>"><%=Guard2HomeNF%>.&nbsp;<%=Guard2HomeN %></a></td>
					</tr>
				</table>
      </div>
		</div>
	</div>
</div>
			<%
				objRSMatchups.MoveNext
				loopcnt = loopcnt + 1
				Wend
			%>
			<%else%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm12 col-xs-12">
			<div class="panel panel-matchups">
				<div class="panel-body">
					<div class="alert alert-info">
						<strong><div style="text-align:center;vertical-align:middle"><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong>&nbsp;NO MATCHUPS SET</br>or</br>NON IGBL GAME DAY!</div>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
			<%end if%>	
</form>
<% 
objRSLineups.Close
objRSLineups.Close
objConn.Close
Set objConn = Nothing
 %>
</body>
</html>
