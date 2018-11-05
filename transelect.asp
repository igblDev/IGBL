<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCI
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""

	Dim sTransaction, sAction, sURL,PID_Split,iRecordToUpdate,strSQL,playerstatus
	Dim objConn, objRS, objChkWaiver, objRS1,errorcode,objRSOwners,objRSrosters,playercnt,chkboxreq,chkfreeagentwaivercnt,objRSTeams
	Dim iPlayerClaimed,iPlayerWaived,objRSplayers,objRSSearch,objRSteam,txnteamname, sMailtext
	Dim email_to, email_subject, host, username, password, reply_to, port, from_address
	Dim ObjSendMail, email_message,objRSCenters,objParams,objRSNBASked,objNextRundate,objrsPlayerPos

	GetAnyParameter  "var_player", sPlayer
	GetAnyParameter  "var_pos", sPosition
	GetAnyParameter "cmbTxnType", sTransaction
	GetAnyParameter "Action", sAction
	GetAnyParameter "var_chkboxreq", chkboxreq
	GetAnyParameter "var_PID", chkPID
	GetAnyParameter "var_sPid", ppPID

	Set objConn		     = Server.CreateObject("ADODB.Connection")
	Set objRS          = Server.CreateObject("ADODB.RecordSet")
	Set objRS1         = Server.CreateObject("ADODB.RecordSet")
	Set objRSOwners    = Server.CreateObject("ADODB.RecordSet")
	Set objRSSearch    = Server.CreateObject("ADODB.RecordSet")
	Set objRSrosters   = Server.CreateObject("ADODB.RecordSet")
	Set objRSplayers   = Server.CreateObject("ADODB.RecordSet")
	Set objChkWaiver   = Server.CreateObject("ADODB.RecordSet")
	Set objRSteam      = Server.CreateObject("ADODB.RecordSet")
	Set objrsPlayerPos = Server.CreateObject("ADODB.RecordSet")
	Set objRSCenters   = Server.CreateObject("ADODB.RecordSet")
	Set objRSForwards  = Server.CreateObject("ADODB.RecordSet")
	Set objRSGuards    = Server.CreateObject("ADODB.RecordSet")
	Set objRSToday     = Server.CreateObject("ADODB.RecordSet")
	Set objParams      = Server.CreateObject("ADODB.RecordSet")
	Set objRSNBASked   = Server.CreateObject("ADODB.RecordSet")
	Set objNextRundate = Server.CreateObject("ADODB.RecordSet")
	Set objTxnAmt      = Server.CreateObject("ADODB.RecordSet")
	Set objRSTop20	   = Server.CreateObject("ADODB.RecordSet")
	Set objEmail	    = Server.CreateObject("ADODB.RecordSet")

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
	"Data Source=lineupstest.mdb;" & _
	"Persist Security Info=False"
	%>
	<!--#include virtual="Common/session.inc"-->
	<%	 
	
	objNextRundate.Open	"SELECT * FROM tblTimedEvents where event = 'pendingwaiversall'", objConn
	nextrundate = objNextRundate.Fields("nextrun_est").Value - 1/24
	objNextRundate.Close
	
	objTxnAmt.Open "SELECT * FROM tblParameterCtl where param_name = 'PICKUP' ",objConn
	wPickupAmt = objTxnAmt.Fields("param_amount").Value
	objTxnAmt.Close
	
	objTxnAmt.Open "SELECT * FROM tblParameterCtl where param_name = 'RENT' ", objConn
	wRentAmt = objTxnAmt.Fields("param_amount").Value
	objTxnAmt.Close	
	
	sPlayer = Request.Form("sPlayer")
	PID_Split = Split(Request.Form("Action"), ";")
	sAction = PID_Split(0)
	
	if sAction = "Go" or sAction = "GoTeam" then
		'DO NOTHING
	else
		PID_Split(0) 'SAction
		PID_Split(1) 'PID
		PID_Split(2) 'Name
		PID_Split(3) 'Player Status
	end if

	select case sAction

	case "Compare"
	

				wcomparePID  = Split(Request.Form("chkPID"), ";")
				wcomparePID2 = Split(Request.Form("chkPIDfreeagents"), ";")
				
				comparePID = comparePID &wcomparePID(0)&","
				comparePID = comparePID &wcomparePID2(0)
				Response.Write "Compare Check Boxes Values	 = "&comparePID&".<br>"

			sURL = "compareTabs.asp"
			AddLinkParameter "var_ownerid", ownerid, sURL
			AddLinkParameter "var_comparePID", comparePID, sURL
			Response.Redirect sURL

		
				
	case "PPTransaction"
		sPosition = "Search"

	case "Sign Free Agent"

	'**********************************************************
	'SIGN FREE AGENT PROCESS
	'**********************************************************
		objRSOwners.Open  "SELECT * FROM tblowners WHERE tblowners.OwnerID = "&ownerid, objConn,3,3,1

		playercnt = objRSOwners.Fields("activeplayercnt").value

		iRecordToUpdate = PID_Split(1)
		playerstatus = PID_Split(3)

		if PID_Split(3) = "Free Agent" then
			'DO NOTHING
		else
			chkfreeagentwaivercnt = chkfreeagentwaivercnt + 1
		end if


		if chkfreeagentwaivercnt > 0 then
			errorcode = "Waiver Selected"
		else
			errorcode = "Display Form"
		end if

	case "Process Free Agents"

	'**********************************************************
	'PROCESS FREE AGENTS
	'**********************************************************
		GetAnyParameter "var_chkboxreq", chkboxreq
		GetAnyParameter "var_PID", chkPID
		GetAnyParameter "var_sPid", ppPID
		checkboxCnt = 0 
		if Request.Form("chkPIDfreeagents").Count > 1 then
			errorcode = "Check Box Violation"
		elseif Request.Form("chkPIDfreeagents").Count = 1 and var_playercnt < 14 then
			errorcode = "Display Form"
		elseif Request.Form("chkPIDfreeagents").Count = 0 and var_playercnt < 14 then
			openSpotAvail = true
			errorcode = "Display Form"
		end if	

	case "Free Agent Confirmation"

	'**********************************************************
	'FREE AGENT CONFIRMATION
	'**********************************************************
		if Request.Form("chkPIDfreeagents").Count = cint(chkboxreq) or cint(chkboxreq) <= 0 then

			errorcode = "Update"
			objRS1.Open "SELECT TeamName,ShortName FROM tblowners WHERE (((tblowners.OwnerID)= " & ownerid & "))", objConn
			txnteamname = objRS1.Fields("ShortName").value	
			objRS1.Close
			
			if Request.Form("chkPIDfreeagents").Count = 0 then
			PID_Split = Split(Request.Form("chkPID"), ";")
			PID_Split(0)
			PID_Split(1)
			iPlayerClaimed = PID_Split(0)
			iPlayerClaimedname = PID_Split(1)

			iPlayerWaivedname = "NOBODY.  OPEN ROSTER SPOT"
			sMailtext = "Fellow IGBLERS:" & vbCrLf & _
			"" & txnteamname & " has signed " & iPlayerClaimedname & "."

			strMsgInfo   = txnteamname
			strMsgPlayerS= "Added: "& iPlayerClaimedname
			strMsgPlayerW = null
			
			objRSrosters.Open  "SELECT * FROM tblplayers WHERE PID = "&iPlayerClaimed, objConn,3,3,1
			
			if objRSrosters.Fields("ownerid").value > 0 then
			   errorcode = "Player Not Available"
			else

			'**********************************************************
			'UPDATE TO PLAYER TABLE
			'**********************************************************

			strSQL = "update tblPlayers SET playerStatus = 'O', OwnerId = " & ownerid & ", clearwaiverdate = null WHERE tblPlayers.PID = " & iPlayerClaimed & ";"
			objConn.Execute strSQL

			'**********************************************************
			'UPDATE TO OWNERS TABLE
			'**********************************************************
			strSQL ="update tblowners SET ActivePlayerCnt = ActivePlayerCnt + 1 WHERE ownerid = " & ownerid & ";"
			objConn.Execute strSQL

			'**********************************************************
			'INSERT INTO TRANSACTION TABLE
			'**********************************************************
			TransType = "Added"
			Cost = wPickupAmt
			strSQL ="insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,transAddPlayerCnt) values ('" &_
			ownerid & "', '" & TransType & "','" &  iPlayerClaimed  & "', '" &  Cost & "', 1)"
			objConn.Execute strSQL
		
			end if

		else

		PID_Split = Split(Request.Form("chkPIDfreeagents"), ";")
		PID_Split(0)
		PID_Split(1)
		iPlayerWaived = PID_Split(0)
		iPlayerWaivedname = PID_Split(1)
		
		objRSrosters.Open  "SELECT * FROM tblplayers WHERE tblplayers.PID = "&iPlayerWaived, objConn,3,3,1
		iPlayerWaivedname = left(objRSrosters.Fields("firstName").Value,1) & ". " & objRSrosters.Fields("lastName").Value 
		objRSrosters.Close
		
		
		PID_Split = Split(Request.Form("chkPID"), ";")
		PID_Split(0)
		PID_Split(1)
		iPlayerClaimed = PID_Split(0)
		iPlayerClaimedname = PID_Split(1)
		
		sMailtext = "Fellow IGBLERS:" & vbCrLf & _
		"" & txnteamname & " has waived " & iPlayerWaivedname & "." & vbCrLf & vbCrLf & _
		"" & txnteamname & " has signed " & iPlayerClaimedname & "."

		objRSrosters.Open  "SELECT * FROM tblplayers WHERE tblplayers.PID = "&iPlayerClaimed, objConn,3,3,1
		iPlayerClaimedname = left(objRSrosters.Fields("firstName").Value,1) & ". " & objRSrosters.Fields("lastName").Value 
		
		strMsgInfo   = txnteamname
		strMsgPlayerS= "Added: "& iPlayerClaimedname
		strMsgPlayerW= "Dropped: "& iPlayerWaivedname
		
		if objRSrosters.Fields("ownerid").value > 0 then
			errorcode = "Player Not Available"
		else

		'**********************************************************
		'UPDATE TO PLAYER TABLE
		'**********************************************************
		strSQL = "update tblPlayers SET playerStatus = 'W', LastTeamInd = "& ownerid &", " &_
		"OwnerId = 0,  clearwaiverdate = Date() + 1 " &_
		"WHERE tblPlayers.PID = " & iPlayerWaived & ";"
		objConn.Execute strSQL
		
		FuncCall = Remove_From_Lineups(iPlayerWaived,0,0,0,0,0,ownerid,0)

		strSQL = "update tblPlayers SET playerStatus = 'O', OwnerId = "& ownerid &", " &_
		"clearwaiverdate = null " &_
		"WHERE tblPlayers.PID = " & iPlayerClaimed & ";"
		objConn.Execute strSQL

		'**********************************************************
		'UPDATE TO OWNERS TABLE
		'**********************************************************

		if iPlayerWaived = 0 then
			strSQL ="update tblowners SET ActivePlayerCnt = ActivePlayerCnt + 1 WHERE ownerid = "&ownerid
		end if

		objConn.Execute strSQL

		'**********************************************************
		'INSERT INTO TRANSACTION TABLE
		'**********************************************************
		'TransType = "Signed free agent"
		TransType = "Added"
		Cost = wPickupAmt
		
        strSQL ="insert into tblTransactions " & _
				"(OwnerID,TransType,TransCost,transAddPlayerCnt,transAddPlayer1,transReleasePlayerCnt,transReleasePlayer1) " & _
		        "values ("&ownerid&",'"&TransType&"',"&Cost&",1,"&iPlayerClaimed&",1,"&iPlayerWaived&") "		
		objConn.Execute strSQL

		end if
		
		end if
		
		
		if errorcode = "Update" then
		'****************************************
		'EMAIL TO THE LEAGUE OF FREE AGENT Pick-up
		'****************************************
		
		wEmailOwnerID       = null
		wAlert              = "receiveFreeAgentAlerts"
		email_subject       = "Free Agent Signed by " & strMsgInfo 
		email_message       = email_message & strMsgPlayerS & "<br>"
		email_message     	= email_message & strMsgPlayerW & "<br>"		
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%		
		end if
		else
			errorCode = "Missing Checkbox"
		end if

	case "Process Rental"

	'**********************************************************
	'RENT PLAYERS CONFIRMATION
	'**********************************************************

	PID_Split = Split(Request.Form("chkPID"), ";")
	PID_Split(0)
	PID_Split(1)
	PID_Split(2)

	iRecordToUpdate = PID_Split(0)
	playerstatus = PID_Split(2)
	'rentalplayer = PID_Split(1)
	
	objRS1.Open "SELECT TeamName,ShortName FROM tblowners WHERE tblowners.OwnerID = "&ownerid, objConn
	txnteamname = objRS1.Fields("ShortName").value	
	objRS1.Close
	
	objRSrosters.Open  "SELECT * FROM tblplayers WHERE PID = "&iRecordToUpdate, objConn,3,3,1
	rentalplayer = left(objRSrosters.Fields("firstName").Value,1) & ". " & objRSrosters.Fields("lastName").Value 
	
	if objRSrosters.Fields("ownerid").value > 0 then
		errorcode = "Player Not Available"
	else
		errorcode = false
		
		wEmailOwnerID = null
		wAlert        = "receiveRentalAlerts"
		email_subject = "Player Rented by " & txnteamname 
		email_message = rentalplayer
%>		
		<!--#include virtual="Common/email_league.inc"-->
				   
<%			

		'**********************************************************
		'INSERT INTO TRANSACTION TABLE
		'**********************************************************
		TransType = "Rental"
		Cost = wRentAmt

		strSQL ="insert into tblTransactions(OwnerID,TransType,transAddPlayer1,TransCost,transAddPlayerCnt) values ('" &_
		ownerid & "', '" & TransType & "','" &  iRecordToUpdate  & "', '" &  Cost & "', 1)"
		
		objConn.Execute strSQL

		'**********************************************************
		'UPDATE TO PLAYER TABLE
		'**********************************************************
		strSQL = "update tblPlayers SET PlayerStatus = 'O', rentalplayer = Yes, OwnerID = "&ownerid&" WHERE tblPlayers.PID = "&iRecordToUpdate
		objConn.Execute strSQL
	end if

	case "Rent Player(s)"

	'**********************************************************
	'RENT PLAYERS
	'**********************************************************

	iRecordToUpdate = PID_Split(1)
	playerstatus = PID_Split(3)
	rentalplayer = PID_Split(2)

	if PID_Split(2) = "Free Agent" then
	'DO NOTHING
	else
		chkfreeagentwaivercnt = chkfreeagentwaivercnt + 1
	end if

	objRSrosters.Open  "SELECT * FROM tblplayers WHERE PID = "&iRecordToUpdate, objConn,3,3,1

	if objRSrosters.Fields("ownerid").value > 0 then
		errorcode = "Player Not Available"
	else
		errorcode = false
	end if

	case "Waiver Claim"

	'**********************************************************
	'WAIVER CLAIM
	'**********************************************************
	objRSOwners.Open  "SELECT * FROM tblowners WHERE tblowners.OwnerID = "&ownerid, objConn,3,3,1
	playercnt = objRSOwners.Fields("activeplayercnt").value

	if Request.Form("chkPIDfreeagents").Count = 0 and playercnt >=14 then
		errorcode = "Check Box Violation"
	end if

	iRecordToUpdate = PID_Split(1)
	playerstatus = PID_Split(3)

	if PID_Split(3) = "Waivers" or PID_Split(3) = "Staggered" then
	'DO NOTHING
	else
		chkfreeagentwaivercnt = chkfreeagentwaivercnt + 1
	end if

	if chkfreeagentwaivercnt > 0 then
		errorcode = "Free Agent Selected"
	else

		objChkWaiver.Open "SELECT * FROM tblPlayers " &_
		"where PID =" & iRecordToUpdate & " and " &_
		"LastTeamInd = " & ownerid & " ;" , objConn,3,3,1

		pName =  objChkWaiver.Fields("firstName").Value & " " & objChkWaiver.Fields("lastName").Value 
		w_count = 	objChkWaiver.Recordcount
		objChkWaiver.Close

		if w_count >= 1 then
			errorcode = "Waiver ineligible"
		else
			chkboxreq = (playercnt + Request.Form("chkPID").Count) - 14
			errorcode = "Display Form"
		end if

	end if


	case "Process Waivers"

	'**********************************************************
	'PROCESS WAIVERS
	'**********************************************************

		if Request.Form("chkPIDfreeagents").Count > 1 then
			errorcode = "Check Box Violation"
		elseif Request.Form("chkPIDfreeagents").Count = 1 and var_playercnt < 14 then
		  errorcode = "Display Form"
		elseif Request.Form("chkPIDfreeagents").Count = 0 and var_playercnt < 14 then
			openSpotAvail = true
			errorcode = "Display Form"
		end if

		pidBidAmount = Request.Form("bidAmount")	
		'response.write "Claimed Player BId = " & pidBidAmount & " <br>" 

case "Waiver Confirmation"

'**********************************************************
'WAIVER CONFIRMATION
'**********************************************************
objRSOwners.Open  "SELECT * FROM tblowners WHERE tblowners.OwnerID = "&ownerid, objConn,3,3,1

playercnt = objRSOwners.Fields("activeplayercnt").value

if Request.Form("chkPIDfreeagents").Count = 0 and playercnt >= 14 then
errorcode = "Missing Checkbox"
else

PID_Split = Split(Request.Form("chkPIDfreeagents"), ";")
PID_Split(0)
PID_Split(1)
iPlayerWaived = PID_Split(0)

PID_Split = Split(Request.Form("chkPID"), ";")
PID_Split(0)
PID_Split(1)
iPlayerClaimed = PID_Split(0)

iwaiverBid = Request.Form("pidBidAmount")

objChkWaiver.Open "SELECT * FROM tblWaivers " & _
                  "WHERE PID_Waived = "&iPlayerWaived&" and PID_Claimed = "&iPlayerClaimed&" and OwnerID = "&ownerid, objConn,3,3,1
w_count = objChkWaiver.Recordcount
objChkWaiver.Close

if w_count >= 1 then
errorcode = "Already Claimed"
else

'**********************************************************
'UPDATE TO PLAYER TABLE
'**********************************************************
strSQL = "update tblPlayers SET pendingwaiver = YES WHERE tblPlayers.PID = "&iPlayerWaived
objConn.Execute strSQL

'**********************************************************
'INSERT INTO PENDING WAIVERS TABLE
'**********************************************************

if Request.Form("chkPIDfreeagents").Count = 0	then
   iPlayerWaived = 0
end if

strSQL ="insert into tblwaivers(OwnerId,PID_Claimed,PID_Waived,waiverbid) " &_
        "values ("&ownerid&","&iPlayerClaimed&","&iPlayerWaived&","&iwaiverBid&")"
objConn.Execute strSQL

strSQL ="insert into tblwaiverlog(OwnerId,PID_Claimed,PID_Waived,waiverbid) " &_
        "values ("&ownerid&","&iPlayerClaimed&","&iPlayerWaived&","&iwaiverBid&")"
objConn.Execute strSQL

errorcode = "Update"

end if

end if

case ""

		ownerid = session("ownerid")	

		if ownerid = "" then
			GetAnyParameter "var_ownerid", ownerid
		end if
		
case "Cancel Rental"

'**********************************************************
'CANCEL BUTTON HIT
'**********************************************************

sURL = "dashboard.asp"
AddLinkParameter "var_ownerid", ownerid, sURL
Response.Redirect sURL


case "Cancel Free Agent Add"

'**********************************************************
'CANCEL BUTTON HIT
'**********************************************************

sURL = "dashboard.asp"
AddLinkParameter "var_ownerid", ownerid, sURL
Response.Redirect sURL

case "Cancel Waiver Request"

'**********************************************************
'CANCEL BUTTON HIT
'**********************************************************

sURL = "dashboard.asp"
AddLinkParameter "var_ownerid", ownerid, sURL
Response.Redirect sURL

end select

%>
<!--#include virtual="Common/functions.inc"-->
<!--#include virtual="Common/setStaggeredAll.inc"-->
<!--#include virtual="Common/setwaiversall.inc"-->
<%
	param_stagger = "STAGGER_WINDOW"
	objParams.Open  "SELECT * FROM tblParameterCtl WHERE param_name = '"&param_stagger&"'", objConn,3,3,1
	staggerPeriod = objParams.Fields("param_indicator").value
	objParams.Close
%>	
<!DOCTYPE HTML>
<html lang="en">
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<link href='//fonts.googleapis.com/css?family=Pattaya' rel='stylesheet'>
<style>

.center {
    text-align: center;
}
.modal-header-success {
    color:white;
    padding:9px 15px;
    border-bottom:1px solid #eee;
    background-color: #354478;
    -webkit-border-top-left-radius: 5px;
    -webkit-border-top-right-radius: 5px;
    -moz-border-radius-topleft: 5px;
    -moz-border-radius-topright: 5px;
     border-top-left-radius: 5px;
     border-top-right-radius: 5px;
		 font-weight:bold;
}
.modal-header-modal {
    color:black;
    padding:9px 15px;
    border-bottom:1px solid #eee;
    background-color: yellowgreen;
    -webkit-border-top-left-radius: 5px;
    -webkit-border-top-right-radius: 5px;
    -moz-border-radius-topleft: 5px;
    -moz-border-radius-topright: 5px;
     border-top-left-radius: 5px;
     border-top-right-radius: 5px;
		 font-weight:bold;
}
.modal-content {
    position: relative;
    background-color: #d9ded1;
    -webkit-background-clip: padding-box;
    background-clip: padding-box;
    border: 1px solid #999;
    border: 1px solid rgba(0,0,0,.2);
    border-radius: 6px;
    outline: 0;
    -webkit-box-shadow: 0 3px 9px rgba(0,0,0,.5);
    box-shadow: 0 3px 9px rgba(0,0,0,.5);
}
.alert-info {
    color: #354478;
    background-color: #f2ff00;
    border-color: #354478;
}
.th {
	white-space:nowrap !important;
}
.td {
	font-size:11px !important;
}
waiver {
	color:#9a1400;
	font-size:10px;	
	font-weight:bold;
  text-transform: uppercase;
}
.bs-callout-success h4 {
    color: black;
}
.bs-callout-success {
    border-left-color: black;
    padding: 10px;
    border-left-width: 4px;
    border-radius: 3px;
    background-color: white;
}

white {
	color:white;
	font-weight: 500;
}
.fa-plus { 
line-height: inherit;
}

table.dataTable,
table.dataTable th,
table.dataTable td {
	-webkit-box-sizing: content-box;
	-moz-box-sizing: content-box;
	box-sizing: content-box;
}

.alert-danger {
    color: #ffffff;
    background-color: #9a1400;
    border-color: #111;
}
dateText {
	color:black;
	font-weight: bold;
}

select {
    color: black;
    background-color: #9a1400;
    color: white;
}
.jumbotron {
    padding-top: 30px;
    padding-bottom: 30px;
    margin-bottom: 30px;
    color: inherit;
    background-color: #d9ded1;
}

.btn-txn-red {
    color: #9a1400;
    background-color: #f9f9f9;
    border-color: #999 !important;
    font-weight: bold !important;
    border: #ffffff;
    padding: 4px 5px;
    line-height: 1.5;
    font-size: 10px;
    border: 2px solid #a1a1a1;
    border-width: thick;
}
.btn-txn-green {
    color:#468847;
    background-color: #f9f9f9;
    border-color: #999 !important;
    font-weight: bold !important;
    border: #ffffff;
    padding: 4px 5px;
    line-height: 1.5;
    font-size: 10px;
    border: 2px solid #a1a1a1;
    border-width: thin !important;
}

.btn-txn-blue {
    color: rgb(1, 87, 155);
    background-color: #f9f9f9;
    border-color: #999 !important;
    font-weight: bold !important;
    border: #ffffff;
    padding: 4px 5px;
    line-height: 1.5;
    font-size: 10px;
    border: 2px solid #a1a1a1;
    border-width: thick;
}
.mark-yellow {
    color: #000;
    background: #ff0;
}
.panel-override {
    background-color: #d9ded1;
    border-color: black;
    border-width: 1px;
}
</style>
</head>
<body>
<script type="text/javascript">
<!--ON SCREEN EDITS-->
function processTransactions(theForm) {
	
	var playerCnt = 0;
	var waiverBid = 0;
	var action = null
	
	playerCnt = theForm.elements["var_playercnt"].value;
	waiverBid = theForm.elements["bidAmount"].value;
	action    = theForm.elements["var_Action"].value;

	var playerSelected = 0;
	for(var i=0; i < theForm.chkPIDfreeagents.length; i++){
		if(theForm.chkPIDfreeagents[i].checked) {
		playerSelected +=1;
		}
	}

	if(playerSelected == 0 && playerCnt >= 14 ) {
		alert("Select Player to Release! " + playerCnt + " Players On Your Roster" ); 
		return (false);
	}

	if(action == "Waiver Claim" ) {
		if(waiverBid == 0 ) {
		alert("Select Waiver Bid Amount!" ); 
		return (false);
		}	
	}
	

	return (true);
}	



$(document).ready(function() {
    $('#example').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false
    } );
} );
$(document).ready(function() {
    $('#example1').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false
    } );
} );
$(document).ready(function() {
    $('#example2').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false
    } );
} );
$(document).ready(function() {
    $('#example3').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false
    } );
} );
$(document).ready(function() {
    $('#example4').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false
    } );
} );
$(document).ready(function() {
    $('#example5').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false
    } );
} );
$(document).ready(function() {
    $('#example6').DataTable( {
        "ordering": false,
				"paging":  	true,
				"info":     false
    } );
} );
$(document).ready(function() {
		$('input[type="checkbox"]').click(function(e) {
			e.preventDefault();
			e.stopPropagation();
		});
	});
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip(); 
});	
</script>
<%
	dim posDisplay

	select case sPosition

	case "Search"
		objRSSearch.Open "SELECT * FROM qry_PlayerAll WHERE qry_PlayerAll.pid = "& (CInt(ppPID)) &" ",objConn,3,3,1
		
		objRSToday.Open   "SELECT * from qry_playerall qp " &_
                      "WHERE PlayerStatus in ('F', 'W', 'S') " &_
											"AND Injury <> 1 " &_
                      "AND exists " &_
                      "(select 1  from nbaindtmsked n " &_
                      "WHERE n.nbateam = qp.nbateamid " &_
											"AND n.gameday = date() ) ", objConn,3,3,1
						
		w_TodayCt = objRSToday.RecordCount
		objRSToday.close
		
	case ""
		objRSToday.Open   "SELECT * from qry_playerall qp " &_
                      "WHERE PlayerStatus in ('F', 'W', 'S') " &_
											"AND Injury <> 1 " &_
                      "AND exists " &_
                      "(select 1  from nbaindtmsked n " &_
                      "WHERE n.nbateam = qp.nbateamid " &_
											"AND n.gameday = date() ) ", objConn,3,3,1
						
		w_TodayCt = objRSToday.RecordCount
		objRSToday.close
 
		loopcnt = 0 
		
	end Select

objRS1.Open "SELECT TeamName,ShortName FROM tblowners WHERE tblowners.OwnerID = "&ownerid, objConn
%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fas fa-user-plus"></i>&nbsp;Player Acquisitions</strong>
			</div>
		</div>
	</div>
</div>
<% if (sAction = "Process Rental" and errorCode = False) then 
	displayAccordian = true 

	Set objRSWork          = Server.CreateObject("ADODB.RecordSet") 
	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'RENT' ",objConn
	wRental = objRSWork.Fields("param_amount").Value
	objRSWork.Close
%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-success">
			<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> Player Rented for Tonight!<br>
					<%
					For I = 1 To Request.Form("chkPID").Count
					PID_Split = Split(Request.Form("chkPID")(I), ";")
					PID_Split(0)
					PID_Split(1)
					rPlayer = PID_Split(1)
					next
					%>
					<strong><%= rPlayer%></strong> has been rented for Tonight's Game. Upon completion of tonight's games <strong><%= rPlayer%></strong> will be available to be claimed off waivers tomorrow.
						<table class="table table-custom-black table-responsive table-condensed" width="100%">
							<tr style="background-color:black;color:white;font-weight:bold;">
								<td style="width:50%;text-align:right;">Transaction Fee:</td><td>$<%=wRental%>.00</td>
							</tr>
						</table>
			</div>
		</div>
	</div>
</div>
<% end if%>
<%if (sAction = "Waiver Claim" or sAction = "Waiver Confirmation") and errorCode = "Already Claimed" then 
displayAccordian = true %>
<form method="POST" name="FrmWaiversErrors4">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_chkboxreq" value="<%= chkboxreq %>" />
<input type="hidden" name="var_playercnt" value="<%= playercnt %>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
			<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i> Waiver Claim Error</strong><br>
				You already have this waiver claim pending. Reasons this may happen are:<br>
				<ul type="bullet">
					<li>
					You hit the browser&#39;s back button 
					</li>
					<li>
					You refreshed your browser after submitting this claim previously
					</li>
				</ul>				
			</div>
		</div>
	</div>
</div>
</form>
<% end if %>
<% if errorCode = "Waiver ineligible" then 
displayAccordian = true
%>
<form method="POST" name="FrmWaiversErrors5">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_chkboxreq" value="<%= chkboxreq %>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
			<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
				<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i> Waiver Claim Error</strong><br>
					You Are Not Eligible to Claim <strong><%= pName %></strong> Off Waivers; You Waived Him.
			</div>
		</div>
	</div>
</div>
</form>
<%end if %>
<% if sAction = "Waiver Claim" and errorCode = "Free Agent Selected" then 
displayAccordian = true
%>
	<div class="container">
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger">
			<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
					<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> Waiver Claim Transaction Error
						<table class="table table-condensed">
						<tr>
							<%
							For I = 1 To Request.Form("chkPID").Count
							PID_Split = Split(Request.Form("chkPID")(I), ";")
							PID_Split(0)
							PID_Split(1)
							%>
							<td align="left"><%=PID_Split(1)%></td>
							<td align="left">Claimed Off Waivers</td>
						</tr>
							<% Next %>
						<tr>
							<%
							For I = 1 To Request.Form("chkPIDfreeagents").Count
							PID_Split = Split(Request.Form("chkPIDfreeagents")(I), ";")
							PID_Split(0)
							PID_Split(1)
							%>
							<td align="left"><%=PID_Split(1)%></td>
							<td align="left">Player Release Pending</td>
							</tr>
							<% Next %>
					</table>					
				</div>
			</div>
		</div>
  </div>
<%end if%>
<% if sAction = "Waiver Confirmation" and errorcode = "Update" then 
	displayAccordian = true
	Set objRSWork = Server.CreateObject("ADODB.RecordSet") 
		
	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PICKUP' ",objConn
	wPickUp = objRSWork.Fields("param_amount").Value
	objRSWork.Close
	Set objRSwaivers       = Server.CreateObject("ADODB.RecordSet")
	objRSwaivers.Open 	"SELECT * FROM tblwaivers WHERE OwnerID = "&ownerid, objConn,3,3,1

 %>
	<div class="container">
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<div class="alert alert-success">
				<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
					<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> Pending Waiver Processed<br>
						<%
						For I = 1 To Request.Form("chkPID").Count
						PID_Split = Split(Request.Form("chkPID")(I), ";")
						PID_Split(0)
						PID_Split(1)
						wPlayer = PID_Split(1)
						next
						
						For I = 1 To Request.Form("chkPIDfreeagents").Count							
						PID_Split = Split(Request.Form("chkPIDfreeagents")(I), ";")
						PID_Split(0)
						PID_Split(1)
						aPlayer = PID_Split(1)
						next
						if aPlayer = "" then
							aPlayer = "Open Roster Spot"
						end if	
						%>
						<% if aPlayer = "Open Roster Spot" then%>
						A waiver claim for <strong><span style="color:white;"><%=wPlayer%></span></strong> has been successfully processed.
						<%else%>
						A waiver claim for <strong><span style="color:white;"><%=wPlayer%></span></strong> has been successfully processed. Upon being awarded <strong><span style="color:white;"><%=aPlayer%></span></strong> will be released and available to be claimed off waivers tomorrow.
						<%end if%>
						</br> 
						<table class="table table-custom-black table-responsive table-bordered table-condensed" width="100%">
							<tr style="background-color:black;color:white;font-weight:bold;">
								<td style="width:50%;text-align:right;">Transaction Fee:</td><td>$<%=wPickUp%>.00</td>
							</tr>
							<tr style="background-color:white;color:#9a1400;font-weight:bold;">
								<td style="width:50%;text-align:right;">View Pending Waivers:</td><td><a href="pendingwaivers.asp?ownerid=<%= ownerid %>" style="color: #9a1400;text-decoration: underline;"><%= objRSwaivers.RecordCount %></a></td>
							</tr>
						</table>
				</div>
			</div>
			
		</div>
  </div>
<%end if%>	
<% if sAction = "Free Agent Confirmation" and errorcode = "Update" then 
	displayAccordian = true 

	Set objRSWork          = Server.CreateObject("ADODB.RecordSet") 
		
	objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PICKUP' ",objConn
	wPickUp = objRSWork.Fields("param_amount").Value
	objRSWork.Close
%>
	<div class="container">
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<div class="alert alert-success">
				<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
					<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong> Free-Agent Pick-up Processed<br>
						<%
						For I = 1 To Request.Form("chkPID").Count
						PID_Split = Split(Request.Form("chkPID")(I), ";")
						PID_Split(0)
						PID_Split(1)
						wPlayer = PID_Split(1)
						next
						
						
						For I = 1 To Request.Form("chkPIDfreeagents").Count							
						PID_Split = Split(Request.Form("chkPIDfreeagents")(I), ";")
						PID_Split(0)
						PID_Split(1)
						aPlayer = PID_Split(1)
						if aPlayer = "" then
							aPlayer = "Open Roster Spot"
						end if	

						next
						%>
						<% if aPlayer = "Open Roster Spot" then%>
							Free-Agent Pickup  <strong><%=wPlayer%></strong> has been successfully processed.
						<%else%>
							Free-Agent Pickup <strong><%=wPlayer%></strong> has been successfully processed. <strong><%=aPlayer%></strong> will be released to waiver pool and available to be claimed tomorrow.
						<%end if%>
						<table class="table table-custom-black table-responsive table-condensed" width="100%">
							<tr style="background-color:black;color:white;font-weight:bold;">
								<td style="width:50%;text-align:right;">Transaction Fee:</td><td>$<%=wPickUp%>.00</td>
							</tr>
						</table>
				</div>
			</div>
		</div>
  </div>
<%end if%>
<!--#include virtual="Common/headerMain.inc"-->	
<% if sAction = "" and sPlayer = "" and sPosition <> "Search" or displayAccordian = true then  %>
<form method="POST" action="transelect.asp" name="frmTransactions">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<div class="container">
  <div class="bs-callout bs-callout-success">
	<h4><span style="font-weight:bold;color:#9a1400;">Transaction Rules</span></h4>
    <ol>
		<li class="big"><span style="text-decoration:underline;font-weight:bold;">Waivers Run:</span>&nbsp;<span style="background-color: yellow;"><dateText><mark><%= (FormatDateTime(nextrundate)) %> cst</span></mark></dateText></li>
		<li class="big">Click on the Position to Display Players</li>
		<li class="big">Select <strong><span class="green">Free-Agent <i class="fa fa-plus green" aria-hidden="true"></i></span> | <span class="blue">Rental</span> <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong> | <span class="red">Waiver <i class="fa fa-plus red" aria-hidden="true"></i></span></strong></li>
		<li class="big"><strong><span><i class="fas fa-user-lock"></i></strong> Not Available to Previous Team via Waivers</span></li>
		<li class="big">Follow Screen Navigation</li>
    </ol>
  </div>
</div>
<br>
<div class="container">
	<div class="row">
		<div class="col-xs-6">
			<button type="button" class="btn btn-trades btn-block " data-toggle="modal" data-target="#myModal"><i class="fa fa-users" aria-hidden="true"></i>&nbsp;Teams in Play</button>
		</div>
		<div class="col-xs-6">
			<button type="button" class="btn btn-trades btn-block " data-toggle="modal" data-target="#myModalWaivers"><i class="fa fa-sort-numeric-asc" aria-hidden="true"></i>&nbsp;Waiver Order</button>
		</div>
	</div>
</div>
<br>
<%if staggerPeriod then %>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger" style="text-align:center;">
				<p><i class="fa fa-refresh fa-spin fa-1x fa-fw margin-bottom"></i> <span class="sr-only"></span> <strong>STAGGER WINDOW IS OPEN!</strong></p>
			</div>
		</div>
	</div>
</div>
<%else%>
<!--<div class="container">
	<div class="bs-callout bs-callout-success">
		<redText><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i>&nbsp;Waivers Next Run:&nbsp;</redText><dateText><%= (FormatDateTime(nextrundate)) %> cst</dateText><br>
	</div>
</div>-->
<%end if%>
<div class="container">
	<div class="row">
		<div id="myModal" class="modal fade" role="dialog">
			<div class="modal-dialog">
				<!-- Modal content-->
				<div class="modal-content">
					<div class="modal-header modal-header-modal">
						<button type="button" class="close" data-dismiss="modal">&times;</button>
						<h5 class="modal-title">Available Teams&nbsp;for&nbsp;<%=date()%><div class="pull-right">all times cst&nbsp;&nbsp;</div></h5>
					</div>
					<% if w_TodayCt > 0 then %>
						<div class="col-md-12 col-sm-12 col-xs-12">
						<br>
						<table class="table table-responsive table-bordered table-condensed">
						<thead>
						<tr class="big" style="background-color:black;color:yellowgreen;font-weight:bold;">
						<td class="big" nowrap  width="50%">Team</td>
						<td class="big" nowrap  width="50%">Game Time</td>
						</tr>
						</thead>
						<% 
						objRSNBASked.Open "SELECT t.teamName, s.NBATeam, " &_ 
						                  "s.GameTime, s.Opponent, s.GameDay " &_
                                          "FROM tblNBATeams t INNER JOIN NBAINDTMSKed s ON t.NBATID = s.NBATeam  " &_
										  "WHERE s.GameDay = date() " &_
										  "order by s.GameTime, t.teamName ", objConn,3,3,1

						While Not objRSNBASked.EOF
						%>
					<tr class="big" style="text-align:left;vertical-align:middle;background-color:white;">
						<td  class="big"><%=objRSNBASked.Fields("teamname").Value %></td>
						<td  class="big"><%=objRSNBASked.Fields("gametime").Value %>	</td>
						</tr>
						<% 
						objRSNBASked.MoveNext
						wend 
						%>
						<%objRSNBASked.close%>
						</table>
					<br>
					</div>
					<div class="modal-footer">
						<button type="button" class=" btn btn-xs btn-trades" data-dismiss="modal">Close</button>
					</div>
					<%else %>
					<div class="jumbotron">
					<h2 style="text-align:center">No Games Today!</h2>
					</div>
					<%end if%>

				</div>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div id="myModalWaivers" class="modal fade" role="dialog">
			<div class="modal-dialog">
				<!-- Modal content-->
				<div class="modal-content">
					<div class="modal-header modal-header-modal">
						<button type="button" class="close" data-dismiss="modal">&times;</button>
						<h5  class="modal-title">Waiver Order</h5>
					</div>
					<div class="col-md-12 col-sm-12 col-xs-12">
						<br>
						<table class="table table-responsive table-bordered table-condensed">
							<thead>
								<tr style="background-color:black;color:yellowgreen;font-weight:bold;">
									<td class="big" width="50%">Team</td>
									<td class="big text-center" width="5%">POS</td>											
									<td class="big text-center" width="5%">BAL</td>									
								</tr>
							</thead>
							<% Set objRSWP = Server.CreateObject("ADODB.RecordSet")	
							   objRSWP.Open "SELECT * FROM tblOwners WHERE ownerID <> 99 AND seasonOver = False " & _
							                "ORDER BY WaiverPriority ", objConn						
							%>
							<%
							 icount = 0
							 While Not objRSWP.EOF
							    icount = icount + 1
							%>
							<%if objRSWP.Fields("ownerid").Value = ownerid then %>
							<tr class="success text-center big" style="font-weight:bold;text-align:left;vertical-align:middle;">
								<td class="big"><%=objRSWP.Fields("TeamName").Value %></td>
								<td class="text-center"><%=objRSWP.Fields("WaiverPriority").Value %></td>
								<td class="big text-center">$<%=objRSWP.Fields("WaiverBal").Value %></td>								
							</tr>
							<%else %>
							<tr class="big" style="text-align:left;vertical-align:middle;background-color:white;">
								<td class="big"><%=objRSWP.Fields("TeamName").Value %></td>
								<td class="big text-center"><%=objRSWP.Fields("WaiverPriority").Value %></td>
								<td class="big text-center">$<%=objRSWP.Fields("WaiverBal").Value %></td>								
							</tr>							
							
							<%end if%>
						<%
							 objRSWP.MoveNext
							 Wend
						%>
						</table>
						</br>
						<table class="table table-responsive table-bordered table-striped table-condensed">
							<tr>
								<td class="big" style="background-color:black;color:yellowgreen;font-weight:bold;" colspan="2" class="text-center">Waiver Rules</td>
							</tr>
							<tr >
								<td class="big;">Period</td>
								<td class="big;">1 Day After Player Waived</td>
							</tr>
							<tr>
								<td class="big">Process</td>
								<td class="big">Owners Making Acquisitions Move to Bottom of Waiver List. <redText><strong>Waivers don't reset.</redText></td>
							</tr>
						</table>
						<br>		
					</div>
					<div class="modal-footer">
						<button type="button" class=" btn btn-xs btn-trades" data-dismiss="modal">Close</button>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<% 
					  
  objrsPlayerPos.Open "SELECT * " &_
                      "FROM qry_playerAll " &_
											"WHERE PlayerStatus in ('F', 'W', 'S') " &_					  
											"AND Injury = FALSE " &_
                      "ORDER BY barps DESC , l5barps DESC , lastname ", objConn,3,3,1
%>
<% if w_seasonOver = true then %>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-danger" style="text-align:center;">
				<p><strong>Your Season is Over | Transactions Not Allowed!</strong></p>
			</div>
		</div>
	</div>
</div>
<%else %>
<div class="container">
  <div class="panel-group" id="accordion">
    <div class="panel panel-override">
      <div class="panel-heading">
        <h5 class="panel-title">
          <a class="accordion-toggle big" data-toggle="collapse" data-parent="#accordion" href="#collapseOne">
            All Players</a>&nbsp;<span class="badgePCnt pull-right"><%= objrsPlayerPos.RecordCount %></span>
        </h5>
      </div>
      <div id="collapseOne" class="panel-collapse collapse">
			<div class="panel-body">
				<table class="table table-custom-black table-responsive table-bordered table-condensed display" style="cellspacing:0;width=:100%" id="example">
					<thead>
						<tr>
							<th style="text-align:center;">Sel</th>
							<th style="text-align:left">Player</th>
							<th style="text-align:center">AVG</th>
							<th style="text-align:center">L/5</th>
						</tr>
					</thead>
					<tbody>
					<%
						While Not objrsPlayerPos.EOF
						
						   wStatus = objrsPlayerPos.Fields("playerStatus").Value
						   wxName = objrsPlayerPos.Fields("lastname").Value
						   
						   objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objrsPlayerPos.Fields("NBATeamID").Value, objConn,3,3,1
						   'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
						   if objRSNBASked.RecordCount > 0 then
						       wTipTime = objRSNBASked.Fields("GameTime").Value							   
							   if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   else
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   end if
						   else
						       wTipTime = "12:00:00 AM"
							   wtime = "12:00 AM"							   
						   end if						   						   							 	
						   
						   'response.write wxName&" "&wStatus&" "&wRent&" <br>" 
						   
						   
				 %>
				 <!--#include virtual="Common/playerPos.inc"-->
					<%
					  objRSNBASked.Close
						loopcnt = loopcnt + 1
						objrsPlayerPos.MoveNext
						Wend
					%>
					</tbody>
				</table>
				</div>
      </div>
    </div>
		<%
		objrsPlayerPos.Close
        objrsPlayerPos.Open "SELECT * " &_
                            "FROM qry_playerAll " &_
														"WHERE PlayerStatus in ('F', 'W', 'S') " &_
														"AND Pos in ('CEN', 'F-C') " &_							
														"AND Injury = FALSE " &_
                            "ORDER BY barps DESC , l5barps DESC , lastname ", objConn,3,3,1							
		%>
    <div class="panel panel-override">
      <div class="panel-heading">
        <h5 class="panel-title">
          <a class="accordion-toggle big" data-toggle="collapse" data-parent="#accordion" href="#collapseTwo">
            Centers</a>&nbsp;<span class="badgePCnt pull-right"><%= objrsPlayerPos.RecordCount %></span>
        </h5>
      </div>
      <div id="collapseTwo" class="panel-collapse collapse">
			<div class="panel-body">
				<table class="table table-striped table-responsive table-custom-black table-bordered table-condensed display" width="100%" id="example1">
					<thead>
						<tr>
							<th style="text-align:center;">Sel</th>
							<th style="text-align:left">Player</th>
							<th style="text-align:center">AVG</th>
							<th style="text-align:center">L/5</th>
						</tr>
					</thead>
					<tbody>
					<%
						While Not objrsPlayerPos.EOF
						   objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objrsPlayerPos.Fields("NBATeamID").Value, objConn,3,3,1
						   
						   'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
						   if objRSNBASked.RecordCount > 0 then
						       wTipTime = objRSNBASked.Fields("GameTime").Value							   
							   if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   else
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   end if
						   else
						       wTipTime = "12:00:00 AM"
							   wtime = "12:00 AM"							   
						   end if						   
						   
						 %>
				 <!--#include virtual="Common/playerPos.inc"-->
					<%
					  objRSNBASked.Close
						objrsPlayerPos.MoveNext
						Wend
					%>					
					</tbody>
				</table>
				</div>
      </div>
    </div>
			<%
		objrsPlayerPos.Close
        objrsPlayerPos.Open "SELECT * " &_
                            "FROM qry_playerAll " &_
														"WHERE PlayerStatus in ('F', 'W', 'S') " &_
														"AND Pos in ('FOR', 'F-C', 'G-F') " &_							
														"AND Injury = FALSE " &_
                            "ORDER BY barps DESC , l5barps DESC , lastname ", objConn,3,3,1									
		%>
    <div class="panel panel-override">
      <div class="panel-heading">
        <h5 class="panel-title">
          <a class="accordion-toggle big" data-toggle="collapse" data-parent="#accordion" href="#collapseThree">
            Forwards</a>&nbsp;<span class="badgePCnt pull-right"><%= objrsPlayerPos.RecordCount %></span>
        </h5>
      </div>
      <div id="collapseThree" class="panel-collapse collapse">
			<div class="panel-body">
				<table class="table table-striped table-responsive table-custom-black table-bordered table-condensed display" width="100%" id="example2">
					<thead>
						<tr>
							<th style="text-align:center;">Sel</th>
							<th style="text-align:left">Player</th>
							<th style="text-align:center">AVG</th>
							<th style="text-align:center">L/5</th>
						</tr>
					</thead>
					<tbody>
					<%
						While Not objrsPlayerPos.EOF
						   objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objrsPlayerPos.Fields("NBATeamID").Value, objConn,3,3,1
						   
						   'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
						   if objRSNBASked.RecordCount > 0 then
						       wTipTime = objRSNBASked.Fields("GameTime").Value							   
							   if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   else
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   end if
						   else
						       wTipTime = "12:00:00 AM"
							   wtime = "12:00 AM"							   
						   end if						   
						 %>
						  <!--#include virtual="Common/playerPos.inc"-->

					<%
						objRSNBASked.Close
						objrsPlayerPos.MoveNext
						Wend
					%>					
					</tbody>
				</table>
				</div>
      </div>
    </div>
		<%
		objrsPlayerPos.Close
		objrsPlayerPos.Open "SELECT * " &_
												"FROM qry_playerAll " &_
												"WHERE PlayerStatus in ('F', 'W', 'S') " &_
												"AND Pos in ('GUA', 'G-F') " &_							
												"AND Injury = FALSE " &_
												"ORDER BY barps DESC , l5barps DESC , lastname ", objConn,3,3,1	
		%>
    <div class="panel panel-override">
      <div class="panel-heading">
        <h5 class="panel-title">
          <a class="accordion-toggle big" data-toggle="collapse" data-parent="#accordion" href="#collapseFour">
            Guards</a>&nbsp;<span class="badgePCnt pull-right"><%= objrsPlayerPos.RecordCount %></span>
        </h5>
      </div>
      <div id="collapseFour" class="panel-collapse collapse">
			<div class="panel-body">
				<table class="table table-striped table-custom-black table-responsive table-bordered table-condensed display" width="100%" id="example3">
					<thead>
						<tr>
							<th style="text-align:center;">Sel</th>
							<th style="text-align:left">Player</th>
							<th style="text-align:center">AVG</th>
							<th style="text-align:center">L/5</th>
						</tr>
					</thead>
					<tbody>
					<%
						While Not objrsPlayerPos.EOF
						   objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objrsPlayerPos.Fields("NBATeamID").Value, objConn,3,3,1
						   
						   'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
						   if objRSNBASked.RecordCount > 0 then
						       wTipTime = objRSNBASked.Fields("GameTime").Value							   
							   if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   else
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   end if
						   else
						       wTipTime = "12:00:00 AM"
							   wtime = "12:00 AM"							   
						   end if						   
					%>
					 <!--#include virtual="Common/playerPos.inc"-->
					<%
						objRSNBASked.Close
						objrsPlayerPos.MoveNext
						Wend
					%>
					</tbody>
				</table>
				</div>
      </div>
    </div>
<!--ROOKIES-->
		<%
		objrsPlayerPos.Close
		objrsPlayerPos.Open "SELECT * " &_
												"FROM qry_playerAll " &_
												"WHERE PlayerStatus in ('F', 'W', 'S') " &_
												"AND Rookie = TRUE " &_							
												"AND Injury = FALSE " &_
												"ORDER BY barps DESC , l5barps DESC , lastname ", objConn,3,3,1	
		
		%>
    <div class="panel panel-override">
      <div class="panel-heading">
        <h5 class="panel-title">
          <a class="accordion-toggle big" data-toggle="collapse" data-parent="#accordion" href="#collapseFive">
            Rookies (Drafted)</a>&nbsp;<span class="badgePCnt pull-right"><%= objrsPlayerPos.RecordCount %></span>
        </h5>
      </div>
      <div id="collapseFive" class="panel-collapse collapse">
			<div class="panel-body">
				<table class="table table-striped table-custom-black table-responsive table-bordered table-condensed display" width="100%" id="example5">
					<thead>
						<tr>
							<th style="text-align:center;">Sel</th>
							<th style="text-align:left">Player</th>
							<th style="text-align:center">AVG</th>
							<th style="text-align:center">L/5</th>
						</tr>
					</thead>
					<tbody>
					<%
						While Not objrsPlayerPos.EOF
						   objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objrsPlayerPos.Fields("NBATeamID").Value, objConn,3,3,1
						   
						   'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
						   if objRSNBASked.RecordCount > 0 then
						       wTipTime = objRSNBASked.Fields("GameTime").Value							   
							   if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   else
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   end if
						   else
						       wTipTime = "12:00:00 AM"
							   wtime = "12:00 AM"							   
						   end if						   
					%>
					 <!--#include virtual="Common/playerPos.inc"-->
					<%
						objRSNBASked.Close
						objrsPlayerPos.MoveNext
						Wend
					%>
					</tbody>
				</table>
				</div>
      </div>
    </div>
<!--END OF ROOKIES-->
		<%
		objrsPlayerPos.Close
		objrsPlayerPos.Open "SELECT * from qry_playerall qp " &_
                            "WHERE PlayerStatus in ('F', 'W', 'S') " &_
														"AND Injury <> 1 " &_		
                            "AND exists " &_												
                            "(select 1  from nbaindtmsked n " &_
														"where n.nbateam = qp.nbateamid " &_
														"AND n.gameday = date() ) " &_
							              "ORDER BY barps DESC , l5barps DESC , lastname ", objConn,3,3,1	
							 
				 
		%>
    <% if objrsPlayerPos.RecordCount > 0 then %>
		<div class="panel panel-override">
      <div class="panel-heading">
        <h5 class="panel-title">
          <a class="accordion-toggle big" data-toggle="collapse" data-parent="#accordion" href="#collapseSix">
            Today</a>&nbsp;<span class="badgePCnt pull-right"><%= objrsPlayerPos.RecordCount %></span>
        </h5>
      </div>
      <div id="collapseSix" class="panel-collapse collapse">
				<div class="panel-body">
				<table class="table table-striped table-custom-black table-responsive table-bordered table-condensed display" width="100%" id="example4">
					<thead>
						<tr>
							<th style="text-align:center;">Sel</th>
							<th style="text-align:left">Player</th>
							<th style="text-align:center">Barps</th>
							<th style="text-align:center">Last 5</th>
						</tr>
					</thead>
					<tbody>
					<%  
						While Not objrsPlayerPos.EOF
						   objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objrsPlayerPos.Fields("NBATeamID").Value &""
						   
						   'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
						   if objRSNBASked.RecordCount > 0 then
						       wTipTime = objRSNBASked.Fields("GameTime").Value							   
							   if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   else
							      wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
							   end if
						   else
						       wTipTime = "12:00:00 AM"
							   wtime = "12:00 AM"							   
						   end if						   
					%>
					 <!--#include virtual="Common/playerPos.inc"-->
					<%
					  objRSNBASked.Close
						objrsPlayerPos.MoveNext
						Wend
					%>
					</tbody>
				</table>
				</div>
      </div>
    </div>
		<% else %>
		<!-- Main jumbotron for a override marketing message or call to action -->
		<div class="panel panel-override">
      <div class="panel-heading">
        <h5 class="panel-title">
          <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion" href="#collapseSix">
            Today </a><span class="badgePCnt pull-right"><%= w_TodayCt %></span>
        </h5>
      </div>
      <div id="collapseSix" class="panel-collapse collapse">
        <div class="panel-body">
					<div class="jumbotron">
						<h2 style="text-align:center">No Game Today!</h2>
							<center>View Free-Agents & Waivers by Position.</center>
					</div>
				</div>
			</div>
	  </div>
	<% end if %>
	</div>
</div>
<%end if %>
</form>
<% end if %>
<% if sPosition  = "Search" then  %>
<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="POST" name="FrmSearch">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<div class="container">
  <div class="row">
  <div class="col-md-12 col-sm-12 col-xs-12">
  <table class="table table-striped table-custom-black table-bordered table-condensed">
    <thead>
			<tr>
				<th style="text-align:center;">Sel</th>
				<th style="text-align:left">Player</th>
				<th style="text-align:center">AVG</th>
				<th style="text-align:center">L/5</th>
			</tr>
    </thead>
    <tbody id="myTable">
    <%
      While Not objRSSearch.EOF
					  
        objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSSearch.Fields("NBATeamID").Value, objConn,3,3,1
     
	    'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.  
        if objRSNBASked.RecordCount > 0 then
	       wTipTime = objRSNBASked.Fields("GameTime").Value							   
	       if len(objRSNBASked.Fields("GameTime").Value) = 10 then
		      wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
		   else
		      wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
		   end if
	    else
	       wTipTime = "12:00:00 AM"
		   wtime = "12:00 AM"							   
	    end if						   
			
		%>
      <tr>			
      <% if objRSSearch.Fields("playerStatus").Value = "W" then %>
        <td style="white-space:nowrap;text-align:center;vertical-align:middle;">
					<button type="submit" value="Waiver Claim;<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>"      name="Action" class="btn justify-content-center align-items-center btn-txn-red btn-xs"><i class="fa fa-plus fa-fw" aria-hidden="true"></i></button>
        </td>
      <% elseif objRSSearch.Fields("playerStatus").Value = "F" and w_TodayCt > 0 then %>
        <% if objRSNBASked.RecordCount > 0 then %>
          <td style="white-space:nowrap;text-align:center;vertical-align:middle;">
  					<button type="submit" value="Sign Free Agent;<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>" name="Action" class="btn justify-content-center align-items-center btn-txn-green btn-xs"><i class="fa fa-plus fa-fw" aria-hidden="true"></i></button>
						<button type="submit" value="Rent Player(s);<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>"  name="Action" class="btn justify-content-center align-items-center btn-txn-blue btn-xs"><i class="fa fa-registered fa-fw" aria-hidden="true"></i></button>
					</td>
        <% else %>
          <td style="white-space:nowrap;text-align:center">
            <button type="submit" value="Sign Free Agent;<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>" name="Action" class="btn btn-freeAgent btn-xs"><i class="fa fa-plus" aria-hidden="true"></i></button>
          </td>
        <% end if %>
       <% elseif objRSSearch.Fields("playerStatus").Value = "S" then %>
        <% if objRSNBASked.RecordCount > 0 AND wTipTime > (time() - 1/24) then %>
          <td style="white-space:nowrap;text-align:center">
						<button type="submit" value="Waiver Claim;<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>"      name="Action" class="btn justify-content-center align-items-center btn-txn-red btn-xs"><i class="fa fa-plus fa-fw" aria-hidden="true"></i></button>
						<button type="submit" value="Rent Player(s);<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>"  name="Action" class="btn justify-content-center align-items-center btn-txn-blue btn-xs"><i class="fa fa-registered fa-fw" aria-hidden="true"></i></button>
          </td>
        <% else %>
          <td style="text-align:center;">
								<button type="submit" value="Waiver Claim;<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>"      name="Action" class="btn justify-content-center align-items-center btn-txn-red btn-xs"><i class="fa fa-plus fa-fw" aria-hidden="true"></i></button>
          </td>
        <% end if %>
      <% else %>
       <td style="white-space:nowrap;text-align:center">
					<button type="submit" value="Sign Free Agent;<%=objRSSearch.Fields("PID").Value & ";" & objRSSearch.Fields("firstName").Value & " " & objRSSearch.Fields("lastName").Value & ";" &  objRSSearch.Fields("StatusDesc").Value%>" name="Action" class="btn justify-content-center align-items-center btn-txn-green btn-xs"><i class="fa fa-plus fa-fw" aria-hidden="true"></i></button>
        </td>
      <% end if %>
      <% if objRSNBASked.RecordCount > 0 then %>
			 <td>
				 <a class="blue" href="playerprofile.asp?pid=<%=objRSSearch.Fields("PID").Value %>" target="_self"><%=objRSSearch.Fields("firstName").Value%>&nbsp;<%=objRSSearch.Fields("lastName").Value%></a>
				 </br><small><span class="greenTrade"><%= objRSSearch.Fields("team").Value%></span>&nbsp;<span class="orange"><%=objRSSearch.Fields("pos").Value%></small></span>
				 </br><span class="gameTip"><%= objRSNBASked.Fields("opponent").Value %>&nbsp;<i class="far fa-clock"></i>&nbsp;<%=wtime%></small></span>
			 </td>
      <% else %>
			<td>
				<a class="blue" href="playerprofile.asp?pid=<%=objRSSearch.Fields("PID").Value %>" target="_self"><%=objRSSearch.Fields("firstName").Value%>&nbsp;<%=objRSSearch.Fields("lastName").Value%></a>
				</br><small><span class="greenTrade"><%= objRSSearch.Fields("team").Value%></span>&nbsp;<span class="orange"><%=objRSSearch.Fields("pos").Value%></small></span>
			</td>
      <% end if %>
			<td class ="big" style="vertical-align:middle;text-align:center;" class="text-center">	
					<span class="badgeBlue big"><%= round(objRSSearch.Fields ("barps").Value,2) %></span>
			</td>
	
			<% if CInt(objRSSearch.Fields("l5barps").Value) > CInt(objRSSearch.Fields("barps").Value) then %>
			<td style="vertical-align:middle;text-align:center;"><span class="badgeUp big"><%= round(objRSSearch.Fields ("l5barps").Value,2) %></span></td>
			<% elseif CInt(objRSSearch.Fields("barps").Value) > CInt(objRSSearch.Fields("l5barps").Value) then%>
			<td style="vertical-align:middle;text-align:center;"><span class="badgeDown big"><%= round(objRSSearch.Fields ("l5barps").Value,2) %></span></td>
			<%else%>
			<td style="vertical-align:middle;text-align:center;"><span class="badgeEven big"><%= round(objRSSearch.Fields ("l5barps").Value,2) %></span></td>
			<%end if %>
			</td>			
		</tr>
		 <%
      objRSNBASked.Close

      objRSSearch.MoveNext
      Wend
      %>
    </table>
  </div>
  </div>
</div>
</form>
<% end if %>
<% if sAction = "Rent Player(s)"  and  errorcode = False then%>
<form method="POST" action="transelect.asp" name="frmwaiverconfirmation">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_PID" value="<%= Request.Form("chkPID")%>" />
<input type="hidden" name="var_chkboxreq" value="<%= chkboxreq %>" />
<input type="hidden" name="var_playercnt" value="<%= playercnt %>" />
<%
objRSrosters.Open   "SELECT * FROM qry_playerall WHERE OwnerID = "&ownerid&" ", objConn,3,3,1
%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
		<div class="panel panel-override">
		<div class="panel-heading" style="background-color:#9a1400;">Rental Confirmation</div>
				<table class="table table-custom-black table-condensed table-bordered">
					<%
					PID_Split(1)
					ppPID = PID_Split(1)
					PID_Split(2)
					PID_Split(3)
					playername = PID_Split(2)
					objRSSearch.Open  "SELECT * FROM qry_PlayerAll WHERE qry_PlayerAll.pid = "& ppPID &" ",objConn,3,3,1

					%>
					<tr>
					  <th><i class="fas fa-basketball-ball" data-toggle="tooltip" title="Playing!"></i></th>
						<th>Name</th>
						<th>Team</th>
					</tr>
					<tr style="background-color:white;">
						<td style="width:5%;" align="center" ><input readonly type="checkbox" class="checkbox" name="chkPID" checked value="<%= PID_Split(1)%>;<%= PID_Split(2)%>;<%= PID_Split(3)%>"></td>
						<td class="text-uppercase red" style="vertical-align:middle;"><%=playername%></td>
						<td class="text-uppercase red" style="vertical-align:middle;"><%=objRSSearch.Fields("teamName").Value %></td>
					</tr>
				</table>
		</div>
		</div>
	</div>
</div>
<% objRSSearch.close %>
<div class="container">
<div class="row">
<div class="col-sm-12 col-md-12" align="right">
<button type="submit" value="Process Rental" name="Action" class="btn btn-danger btn-block"><i class="far fa-arrow-alt-circle-down"></i>&nbsp;Confirm Rental</button>
<!--<button type="submit" onClick="javascript:history.back(-1)"  value="Cancel Rental" name="Action" class="btn btn-trades "><span class="glyphicon glyphicon-trash"></span>&nbsp;Cancel</button>-->
</div>
</div>
</div>
<br>
</form>
<%
end if
if sAction = "Rent Player(s)" and errorCode = "Waiver Selected" then %>
<form method="POST" name="frmRentalError1" language="JavaScript">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_PID" value="<%= Request.Form("chkPID")%>" />
<input type="hidden" name="var_chkboxreq" value="<%= chkboxreq %>" />
<input type="hidden" name="var_playercnt" value="<%= playercnt %>" />
<div class="container">
<div class="row">
<div class="col-md-12 col-sm-12 col-xs-12">
<div class="panel panel-danger">
<div class="panel-heading clearfix"> <i class="icon-calendar"></i>
<h5 class="panel-title">Rental Transaction Error</h5>
</div>
<div class="panel-body">
<table class="table">
<tr>
<td align="left">Name </td>
<td align="left">Status </td>
</tr>
<tr>
<%
For I = 1 To Request.Form("chkPID").Count
PID_Split = Split(Request.Form("chkPID")(I), ";")
PID_Split(0)
PID_Split(1)
PID_Split(2)

%>
<td><input name="txtPlayer" type="hidden" value="<%=PID_Split(1)%> ">
<%=PID_Split(1)%>&nbsp; </td>
<td>Not Eligible to be rented</td>
</tr>
<% Next %>
</table>
<table style="border-collapse:collapse;width:575;" bordercolor="#111111" cellpadding="0" cellspacing="0" border="0">
<tr>
<td style="text-align:center;width:100%;"><br>
<br>
&nbsp;</td>
</tr>
</table>
</div>
</div>
</div>
</div>
</div>
</form>
<%
end if
if sAction = "Sign Free Agent" and errorCode = "Players Signed" then %>
<form method="POST" name="frmfreeagent1" language="JavaScript">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_PID" value="<%= Request.Form("chkPID")%>" />
<input type="hidden" name="var_chkboxreq" value="<%= chkboxreq %>" />
<input type="hidden" name="var_playercnt" value="<%= playercnt %>" />
<div class="container">
<div class="row">
<div class="col-md-12 col-sm-12 col-xs-12">
<div class="panel panel-success">
<div class="panel-heading clearfix">
<h5 class="panel-title">Free Agent Signed</h5>
</div>
<div class="panel-body">
<table class="table table-condensed table-bordered">
<tr>
<td  align="left">Name</td>
<td  align="left">Status</td>
</tr>
<tr>
<%
For I = 1 To Request.Form("chkPID").Count
PID_Split = Split(Request.Form("chkPID")(I), ";")
PID_Split(0)
PID_Split(1)
PID_Split(2)
%>
<td align="center"><input type="hidden" name="txtPlayer" value="<%=PID_Split(1)%> " style="float: left">
<%=PID_Split(1)%>
<p align="left">&nbsp;</p></td>
<td align="center">Signed</td>
</tr>
<% Next %>
</table>
</div>
</div>
</div>
</div>
</div>
</form>
<%
end if

if (sAction = "Waiver Claim" or sAction = "Sign Free Agent") and errorCode <> "Waiver ineligible" then %>
<form action="transelect.asp" method="POST" onSubmit="return processTransactions(this)">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_PID" value="<%= Request.Form("chkPID")%>" />
<input type="hidden" name="var_chkboxreq" value="<%= chkboxreq %>" />
<input type="hidden" name="var_playercnt" value="<%= playercnt %>" />
<input type="hidden" name="var_Action" value="<%= sAction %>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<table class="table table-custom-black table-responsive table-condensed table-bordered">
				<%
				PID_Split(1)
				PID_Split(2)
				PID_Split(3)
				objRSrosters.Open   "SELECT * FROM qry_playerAll WHERE PID = " & PID_Split(1) & " ", objConn,3,3,1
				%>
				<tr>
					<th class="big" colspan="4">Player to Add</th>
				</tr>
				<tr style="background-color:#d9ded1;font-weight:bold;color:grey;vertical-align:middle;">
					<th class="big" style="text-align:center;">Sel</th>
					<th class="big" style="text-align:left;width:80%;">Player</th>
					<th class="big" style="text-align:center;">Avg</th>				
					<th class="big" style="text-align:center;">Bid</th>					
				</tr>
				<tr>
					<td class="big"  align="center" style="background-color:white;width:10%;"><input readonly checked  class="checkbox" type="checkbox" name="chkPID"  value="<%= PID_Split(1)%>;<%= PID_Split(2)%>;<%= PID_Split(3)%>"></td>
					<td class="big" style="background-color:white">
						<a class="blue" href="playerprofile.asp?pid=<%= PID_Split(1)%>" target="_self"><%= PID_Split(2)%></a>,&nbsp;<span class="greenTrade"><%=objRSrosters.Fields("teamshortname").Value %></span>&nbsp;<span class="orangeText"><%=objRSrosters.Fields("pos").Value %></span>
					</td>
					<td class="big" style="vertical-align:middle;text-align:center;background-color:white;width:10%;"><%= round(objRSrosters.Fields("barps").Value,2) %></td>
					<%if sAction = "Waiver Claim" then %>
					<td>        
						<select name="bidAmount" >
						<option value="" selected>Select Bid</option>
							<%
							bidAmountCnt = 1
							While bidAmountCnt <= w_WaiverBal
							%>
							<option value="<%=bidAmountCnt%>"><%=bidAmountCnt %></option>
							<%
								bidAmountCnt = bidAmountCnt +1
								Wend
							%>								
						</select>
						</td>
						<%else%>
						<td class="big red" style="text-align:center;background-color:white;font-weight:bold" >N/A</td>
						<%end if%>
					<%
				'Next
				%>
				</tr>
			</table>
		<br>
		</div>
		<%
		objRSrosters.Close
		objRSrosters.Open   "SELECT * FROM qry_playerAll WHERE OwnerID = "&ownerid&" AND rentalPlayer = 0 " &_
		                    "order by barps desc", objConn,3,3,1
		%>
		<div class="col-md-12 col-sm-12 col-xs-12">
			<table class="table table-custom-black table-condensed table-responsive table-bordered">
				<tr  >
					<th class="big" colspan="10"><span class="pull-left">Player to Drop</span><span class="pull-right"><small>Player Count: <%= playercnt %></small></span></th>
				</tr>
					<tr class="big">
					<th class="big" style="text-align:center;width:5%;vertical-align: inherit;">Sel</th>
					<th class="big" style="text-align:left;width:20%;">Player</th>
					<th class="big hidden-xs" style="vertical-align:inherit;text-align:center;width:10%;" nowrap><span style="color:black;">B</span>/pg</th>
					<th class="big hidden-xs" style="vertical-align:inherit;text-align:center;width:10%;" nowrap><span style="color:black;">A</span>/pg</th>
					<th class="big hidden-xs" style="vertical-align:inherit;text-align:center;width:10%;" nowrap><span style="color:black;">R</span>/pg</th>
					<th class="big hidden-xs" style="vertical-align:inherit;text-align:center;width:10%;" nowrap><span style="color:black;">P</span>/pg</th>
					<th class="big hidden-xs" style="vertical-align:inherit;text-align:center;width:10%;" nowrap><span style="color:black;">S</span>/pg</th>
					<th class="big hidden-xs" style="vertical-align:inherit;text-align:center;width:10%;" nowrap><span style="color:black;">3</span>/pg</th>
					<th class="big hidden-xs" style="vertical-align:inherit;text-align:center;width:10%;" nowrap><span style="color:black;">T</span>/pg</th>
					<th class="big" style="vertical-align:inherit;text-align:center;width:5%;">Avg</th>
				</tr>
				<%
				While Not objRSrosters.EOF
				%>
				<tr style="background-color:white;vertical-align:middle;">
				<% if sAction = "Sign Free Agent" and ((objRSrosters.Fields("pendingTrade")= True) or (objRSrosters.Fields("pendingWaiver")= True)) then%>
					<td class="big" style="width:5%;text-align:center;vertical-align:middle;">
						<redIcon><span data-toggle="tooltip" title="Player Not Available to be Waived!"><i class="fa fa-lock fa-1x"></i></span></redIcon>
					</td>
				<% else %>
					<td class="big" style="text-align:center;vertical-align:middle;"><input type="radio" name="chkPIDfreeagents" value="<%=objRSrosters.Fields("PID").Value & ";" & objRSrosters.Fields("firstName").Value & " " & objRSrosters.Fields("lastName").Value%>"></td>
				<% end if %>		
					<td class="big" style="vertical-align:middle;text-align:left;">
						<%if (len(objRSrosters.Fields("firstName").Value) + len(objRSrosters.Fields("lastName").Value)) >= 17 then %>
							<a class="blue" href="playerprofile.asp?pid=<%=objRSrosters.Fields("PID").Value %>">
								<%=left(objRSrosters.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSrosters.Fields("lastName").Value,12)%></a>,&nbsp;<span class="greenTrade"><%=objRSrosters.Fields("teamshortname").Value %></span>&nbsp;<span class="orangeText"><%=objRSrosters.Fields("pos").Value %></span>
						<%else%>
							<a class="blue" href="playerprofile.asp?pid=<%=objRSrosters.Fields("PID").Value %>">
							<%=objRSrosters.Fields("firstName").Value%>&nbsp;<%=objRSrosters.Fields("lastName").Value%></a>,&nbsp;<span class="greenTrade"><%=objRSrosters.Fields("teamshortname").Value %></span>&nbsp;<span class="orangeText"><%=objRSrosters.Fields("pos").Value %></span>
						<%end if%>						
					</td>
					<td class="hidden-xs big" style="text-align:center;vertical-align:middle;"><%=objRSrosters.Fields("blk").Value %></td>
					<td class="hidden-xs big" style="text-align:center;vertical-align:middle;"><%=objRSrosters.Fields("ast").Value %></td>
					<td class="hidden-xs big" style="text-align:center;vertical-align:middle;"><%=objRSrosters.Fields("reb").Value %></td>
					<td class="hidden-xs big" style="text-align:center;vertical-align:middle;"><%=objRSrosters.Fields("ppg").Value %></td>
					<td class="hidden-xs big" style="text-align:center;vertical-align:middle;"><%=objRSrosters.Fields("stl").Value %></td>
					<td class="hidden-xs big" style="text-align:center;vertical-align:middle;"><%=objRSrosters.Fields("three").Value %></td>
					<td class="hidden-xs big" style="text-align:center;vertical-align:middle;"><%=objRSrosters.Fields("to").Value %></td>
					<td class="big" style="vertical-align:middle;text-align:center;"><%= round(objRSrosters.Fields("barps").Value,2) %></td>
				</tr>
				<%
				objRSrosters.MoveNext
				Wend
				%>
			</table>
		<br>
		</div>
	</div>
</div>
<% if sAction = "Waiver Claim" then %>
<div class="container">
	<div class="row">
		<div class="col-xs-6">
			<%if errorCode = "Waiver ineligible" then %>
			<button type="submit" disabled value="Process Waivers" name="Action" class="btn  btn-block btn-default  "><i class="fa fa-step-forward" aria-hidden="true"></i>&nbsp;Next</button>
			<%else%>
			<button type="submit" value="Process Waivers" name="Action" class="btn btn-default btn-block "><i class="fa fa-step-forward" aria-hidden="true"></i>&nbsp;Next</button>
			<%end if%>
		</div>
		<div class="col-xs-6">
				<button type="reset"  class="btn btn-block btn-default "><i class="fa fa-refresh" aria-hidden="true"></i>&nbsp;Reset</button>
		</div>
	</div>
</div>
<%else%>
<div class="container">
	<div class="row">
		<div class="col-xs-6">
		<button type="submit" value="Process Free Agents" name="Action" class="btn btn-block btn-default "><i class="fa fa-step-forward" aria-hidden="true"></i>&nbsp;Next</button>
  	<!--<button type="submit" onClick="javascript:history.back(-1)" value="Cancel Free Agent Add" name="Action" class="btn btn-trades "><span class="glyphicon glyphicon-trash"></span>&nbsp;Cancel</button>-->
		</div>
		<div class="col-xs-6">
				<button type="reset"  class="btn btn-block btn-default "><i class="fa fa-refresh" aria-hidden="true"></i>&nbsp;Reset</button>
		</div>
	</div>
</div>
<%end if%>
<br>
</form>
<%
end if
if (sAction = "Process Waivers"  or sAction = "Process Free Agents") and  errorcode = "Display Form" then
%>
<form method="POST" action="transelect.asp" name="frmwaiverconfirmation" language="JavaScript">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<input type="hidden" name="var_PID" value="<%= Request.Form("chkPID")%>" />
<input type="hidden" name="var_chkboxreq" value="<%= chkboxreq %>" />
<input type="hidden" name="var_playercnt" value="<%= playercnt %>" />
<input type="hidden" name="pidBidAmount" value="<%= pidBidAmount %>" />
<%
objRSrosters.Open   "SELECT * FROM qry_playerAll WHERE OwnerID = "&ownerid&" ", objConn,3,3,1
%>
<% if sAction = "Process Waivers" then %>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<table class="table table-custom-black table-bordered table-responsive table-condensed">
				<tr class="text-center">
					<th width="50%">Bid Amount</th>
					<th>Action</th>
				</tr>
				<tr class="text-center" style="background-color:white;">
					<td style="background-color: yellowgreen;color: white;font-weight:bold;"><mark><%= FormatCurrency(pidBidAmount)%></mark></td>
					<td style="background-color: yellowgreen;color: white;font-weight:bold;"><mark>Submit Waivers Request</mark></td>		
				</tr>	
			</table>
		</div>
	</div>		
</div>
</div>
</br>
<%end if%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<table class="table table-responsive table-custom-black table-condensed table-bordered">
	
				<tr class="text-center">
					<th class="big" width="50%">Adding</th>
					<th class="big" style="color:#9a1400;width:50%">Dropping</th>
				</tr>
				<tr style="background-color:#d9ded1;">
				<%
						For I = 1 To Request.Form("chkPID").Count
						PID_Split = Split(Request.Form("chkPID")(I), ";")
						PID_Split(0)
						PID_Split(1)
						PID_Split(2)
						playername = PID_Split(1)
				%>
					<td class="big" width="50%" align="center"><input readonly type="checkbox" class="checkbox" name="chkPID" checked value="<%= Request.Form("chkPID") (I)%>"></td>
				<%
					Next
				%>				<% if openSpotAvail then %>
					<td  class="big" width="50%" align="center"><input class="checkbox" readonly type="checkbox" name="0"></td>
				<%else %>
				<%
				
					For I = 1 To Request.Form("chkPIDfreeagents").Count
						PID_Split2 = Split(Request.Form("chkPIDfreeagents")(I), ";")
						PID_Split2(0)
						PID_Split2(1)
						PID_Split2(2)
						playername = PID_Split2(1)
				%>
					<td class="big" width="50%" align="center"><input class="checkbox" readonly type="checkbox" name="chkPIDfreeagents" checked value="<%= Request.Form("chkPIDfreeagents") (I)%>"></td>
				<%
					Next
				%>
				<% end if%>


				<%
					objRS.Close
					objRS1.Close
					
					objRS.Open  "SELECT * FROM qry_playerAll WHERE pid = "&PID_Split2(0) ,objConn,1,1
					objRS1.Open "SELECT * FROM qry_playerAll WHERE pid = "&PID_Split(0) ,objConn,1,1
				%>
				</tr>
				<tr style="background-color:white">
					<td   class="big" width="50%" align="center"><img class="img-responsive img-circle center-block" width="260px" height="190px" src="<%=objRS1.Fields("image").Value%>"></td>
				<% if openSpotAvail then %>
						<td class="big" width="50%" align="center"><img class="img-responsive img-circle center-block" width="260px" height="190px" src="http://i.cdn.turner.com/nba/nba/.element/img/2.0/sect/statscube/players/large/default_nba_headshot_v2.png"></td>
					<% else %>
						<td class="big" width="50%" align="center"><img class="img-responsive img-circle center-block" width="260px" height="190px" src="<%=objRS.Fields("image").Value%>"></td>
					<% end if %>
				</tr>
				<tr style="background-color:#d9ded1;color:white;font-weight:bold;">
				<td class="big text-uppercase" align="center"><a class="blue" href="playerprofile.asp?pid=<%=objRS1.Fields("PID").Value %>"><%=left(objRS1.Fields("firstName").Value,14)%>&nbsp;<%=objRS1.Fields("lastName").Value%></a></td>
				<% if openSpotAvail then %>
					<td class="big blue" align="center">OPEN SPOT</td>
				<%else%>
					<td class="big text-uppercase" align="center"><a class="red" href="playerprofile.asp?pid=<%=objRS.Fields("PID").Value %>"><%=objRS.Fields("firstName").Value%>&nbsp;<%=objRS.Fields("lastName").Value%></a></td>
				<%end if%>
				</tr>
				<tr style="background-color:white;">
					<td class="big" style="text-align:center;vertical-align:middle;"><span class="greenTrade"><%=objRS1.Fields("teamname").Value %></span>&nbsp;<span class="orange"><%=objRS1.Fields("pos").Value %></span></td>
					<% if openSpotAvail then %>
					<td style="text-align:center;vertical-align:middle;" >N/A</td>
					<%else%>
						<td class="big" style="text-align:center;vertical-align:middle;"><span class="red"><%=objRS.Fields("teamname").Value %></span>&nbsp;<span class="red"><%=objRS.Fields("pos").Value %></span>	</td>
				
					<%end if%>
				</tr>
			</table>
		</div>
	</div>
</div>
<br>
<% if sAction = "Process Free Agents" then %>
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
			<button type="submit" value="Free Agent Confirmation" name="Action" class="btn btn-danger btn-block "><i class="far fa-arrow-alt-circle-down"></i>&nbsp;Confirm Free-Agent P/U</button>
			<% if openSpotAvail = false then %>
			  <button type="button" class="btn btn-default btn-block " data-toggle="modal" data-target="#compareModal"><i class="fa fa-balance-scale" aria-hidden="true"></i>&nbsp;Compare Players</button>
				<%end if%>
		</div>
	</div>
</div>
<% else %>
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
			<button type="submit" value="Waiver Confirmation" name="Action" class="btn btn-danger btn-block "><i class="far fa-arrow-alt-circle-down"></i>&nbsp;Confirm Waiver Request</button>
			<% if openSpotAvail = false then %>
			  <button type="button" class="btn btn-default btn-block " data-toggle="modal" data-target="#compareModal"><i class="fa fa-balance-scale" aria-hidden="true"></i>&nbsp;Compare Players</button>
			<%end if%>		
		</div>
	</div>
</div>
<% end if %>
</form>
<br>
<%
end if
%>
<div class="container">
	<div class="row">
		<div id="compareModal" class="modal fade" role="dialog">
			<div class="modal-dialog">
				<!-- Modal content-->
				<div class="modal-content">
					<div class="modal-header modal-header-modal">
						<button type="button" class="close" data-dismiss="modal">&times;</button>
						<h4 class="modal-title">LAST 5 PLAYER COMPARISON</h4>
					</div>
						<%
							Set objRSAll  = Server.CreateObject("ADODB.RecordSet")	
							Set objsName  = Server.CreateObject("ADODB.RecordSet")	
							Set objsBarps = Server.CreateObject("ADODB.RecordSet")								
							
							<!--FIRST PLAYER SELECTED-->
							objsName.Open "Select firstName,lastName,POS,PID,image  from tblPlayers where PID = "&PID_Split(0)&" ", objConn,3,3,1
							
							wFirstName    = objsName.Fields("firstName").Value
							wLastName     = objsName.Fields("lastName").Value
							intFName      = left(wFirstName,2)
							intLName			= left(wLastName,2)
							image         = objsName.Fields("image").Value
							wPos          = objsName.Fields("POS").Value

							objsBarps.Open "SELECT barps,usage FROM tbl_barps t WHERE t.first = '" & wFirstName & "'  and t.last ='"&wLastName & "' ", objConn,1,1					
							
							objRSAll.Open "SELECT avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
														"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP " &_
														"FROM tblLast5 t " & _
														"WHERE t.first = '" & wFirstName & "'  and t.last ='"&wLastName & "' ", objConn,1,1
					
							avgMP    = objRSAll.Fields("avgMP").Value
							avgBlks  = objRSAll.Fields("avgBlks").Value
							avgAst   = objRSAll.Fields("avgAst").Value
							avgReb   = objRSAll.Fields("avgReb").Value
							avgPts   = objRSAll.Fields("avgPts").Value
							avgStl   = objRSAll.Fields("avgStl").Value
							avg3pt   = objRSAll.Fields("avg3pt").Value
							avgTo    = objRSAll.Fields("avgTo").Value
							avgBarps = objRSAll.Fields("avgBarps").Value
							barps    = objsBarps.Fields("Barps").Value 
							usage    = objsBarps.Fields("usage").Value 
							
							objRSAll.close	
							objsName.close	
							objsBarps.close		
						<!--EOF FIRST PLAYER SELECTED-->	
						<!--SECOND PLAYER SELECTED-->							
							objsName.Open "Select firstName,lastName,POS,PID,image  from tblPlayers where PID = "&PID_Split2(0)&" ", objConn,3,3,1
									
							wFirstName2   = objsName.Fields("firstName").Value
							wLastName2    = objsName.Fields("lastName").Value
							intFName2     = left(wFirstName2,2)
							intLName2			= left(wLastName2,2)
							image2        = objsName.Fields("image").Value
							wPos2         = objsName.Fields("POS").Value
							
							
							objsBarps.Open "SELECT barps,usage FROM tbl_barps t WHERE t.first = '" & wFirstName2 & "'  and t.last ='"&wLastName2 & "' ", objConn,1,1					

							objRSAll.Open "SELECT avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
														"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP " &_
														"FROM tblLast5 t " & _
														"WHERE t.first = '" & wFirstName2 & "'  and t.last ='"&wLastName2 & "' ", objConn,1,1
					
							avgMP2    = objRSAll.Fields("avgMP").Value
							avgBlks2  = objRSAll.Fields("avgBlks").Value
							avgAst2   = objRSAll.Fields("avgAst").Value
							avgReb2   = objRSAll.Fields("avgReb").Value
							avgPts2   = objRSAll.Fields("avgPts").Value
							avgStl2   = objRSAll.Fields("avgStl").Value
							avg3pt2   = objRSAll.Fields("avg3pt").Value
							avgTo2    = objRSAll.Fields("avgTo").Value
							avgBarps2 = objRSAll.Fields("avgBarps").Value	
							barps2    = objsBarps.Fields("Barps").Value 	
							usage2    = objsBarps.Fields("usage").Value 		
							
							objRSAll.close	
							objsName.close	
							objsBarps.close		
						<!--EOF SECOND PLAYER SELECTED-->	
						%>
						<table class="table table-custom-black table-responsive table-bordered">	
							<body>
								<tr style="background-color:white;vertical-align:middle;text-align:center;">
									<th class="big">NAME</th>
									<th style="width:45%" class="big"><%=left(wFirstName,1) %>.&nbsp;<%=wLastName %></th>	
									<th style="width:45%" class="big"><%=left(wFirstName2,1) %>.&nbsp;<%=wLastName2 %></th>
								</tr>
								<tr style="background-color:white;vertical-align:middle;text-align:center;">
									<th class="big">POS</th>
									<th class="big"><%=wPos%></th>	
									<th class="big"><%=wPos2%></th>
								</tr>								
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">MPG</th>
									<% if avgMP > avgMP2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgMP,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgMP,2)%></td>								
									<%end if %>
									
									<% if avgMP2 > avgMP then %>	
										<th class="big" style="color:#468847;"><%=round(avgMP2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgMP2,2)%></td>								
									<%end if %>						
								</tr>						
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">BPG</th>	
									<% if avgBlks > avgBlks2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgBlks,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgBlks,2)%></td>								
									<%end if %>
									<% if avgBlks2 > avgBlks then %>	
										<th class="big" style="color:#468847;"><%=round(avgBlks2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgBlks2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">APG</th>	
									<% if avgAst > avgAst2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgAst,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgAst,2)%></td>								
									<%end if %>
									
									<% if avgAst2 > avgAst then %>	
										<th class="big" style="color:#468847;"><%=round(avgAst2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgAst2,2)%></td>								
									<%end if %>		
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">RPG</th>	
									<% if avgReb > avgReb2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgReb,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgReb,2)%></td>								
									<%end if %>
									
									<% if avgReb2 > avgReb then %>	
										<th class="big" style="color:#468847;"><%=round(avgReb2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgReb2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">PPG</th>	
									<% if avgPts > avgPts2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgPts,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgPts,2)%></td>								
									<%end if %>
									<% if avgPts2 > avgPts then %>	
										<th class="big" style="color:#468847;"><%=round(avgPts2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgPts2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">SPG</th>	
									<% if avgStl > avgStl2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgStl,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgStl,2)%></td>								
									<%end if %>
									<% if avgStl2 > avgStl then %>	
										<th class="big" style="color:#468847;"><%=round(avgStl2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgStl2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">3PG</th>	
									<% if avg3pt > avg3pt2 then %>	
										<th class="big" style="color:#468847;"><%=round(avg3pt,2)%></th>	
									<%else%>
										<td class="big"><%=round(avg3pt,2)%></td>								
									<%end if %>
									<% if avg3pt2 > avg3pt then %>	
										<th class="big" style="color:#468847;"><%=round(avg3pt2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avg3pt2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">TPG</th>	
									<% if avgTo < avgTo2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgTo,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgTo,2)%></td>								
									<%end if %>
									<% if avgTo2 < avgTo then %>	
										<th class="big" style="color:#468847;"><%=round(avgTo2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgTo2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">L/5</th>	
									<% if avgBarps > avgBarps2 then %>	
										<th class="big" style="color:#468847;"><%=round(avgBarps,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgBarps,2)%></td>								
									<%end if %>
									<% if avgBarps2 > avgBarps then %>	
										<th class="big" style="color:#468847;"><%=round(avgBarps2,2)%></th>	
									<%else%>
										<td class="big"><%=round(avgBarps2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">Usage</th>	
									<% if usage > usage2 then %>	
										<th class="big" style="color:#468847;"><%=round(usage,2)%></th>	
									<%else%>
										<td class="big"><%=round(usage,2)%></td>								
									<%end if %>
									<% if usage2 > usage  then %>	
										<th class="big" style="color:#468847;"><%=round(usage2,2)%></th>	
									<%else%>
										<td class="big"><%=round(usage2,2)%></td>								
									<%end if %>
								</tr>	
								<tr style="background-vertical-align:middle;text-align:center;background-color:white;">
									<th class="big">Barps</th>	
									<% if barps > barps2 then %>	
										<th class="big" style="color:#468847;"><%=round(barps,2)%></th>	
									<%else%>
										<td class="big"><%=round(barps,2)%></td>								
									<%end if %>
									<% if barps2 > barps then %>	
										<th class="big" style="color:#468847;"><%=round(barps2,2)%></th>	
									<%else%>
										<td class="big"><%=round(barps2,2)%></td>								
									<%end if %>
								</tr>	
							</tbody>
					</table>
					</br>
				</div>
			</div>	
		</div>
	</div>
</div>
<%
objRS.Close
objRSrosters.Close
objRSSearch.Close
ObjConn.Close
objRS1.Close
objRSToday.Close
objRSNBASked.Close
Set objRS = Nothing
Set objConn = Nothing
Set objRSrosters = Nothing
Set objRSSearch = Nothing
Set objRS1 = Nothing
Set objRSNBASked = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>