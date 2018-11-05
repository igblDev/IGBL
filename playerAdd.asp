<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear
	

	Dim objConn,objRSPlayers,strSQL,firstName,lastName,wNextVal,nbaTeam,position,playerStatus,confirmation,image,objRSPIDCnt 
	
	GetAnyParameter "Action", sAction		
	
   	
	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")
  
	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

%>	
	<!--#include virtual="Common/session.inc"-->	
<%								
	Set objRSPIDCnt  = Server.CreateObject("ADODB.RecordSet")
	Set objRSPlayers = Server.CreateObject("ADODB.RecordSet")
	
	objRSPIDCnt.Open "SELECT MAX (pid) +1 as nextval FROM tblPlayers  where pid < 5000", objConn
	
	objRSPlayers.Open	"SELECT tbl_barps.team,tbl_barps.first, tbl_barps.last, tblPlayers.firstName, tblPlayers.lastName " & _
										"FROM tbl_barps LEFT JOIN tblPlayers ON (tbl_barps.last = tblPlayers.lastName) AND (tbl_barps.first = tblPlayers.firstName) " & _
										"WHERE  lastName is null order by tbl_barps.first asc", objConn
        		  
	if Request.Form("action") <> "Save Form Data" Then
		'Nothing
	else
		barpPlayerName= Split(Request.Form("nbaplayername"), ";") 
		firstName     = Request.Form("firstName")
		lastName      = Request.Form("lastName")
		nbaTeam       = Request.Form("teamName")
		position      = Request.Form("position")
		playerStatus  = Request.Form("playerStatus")
		playerCnters  = 0
		wNextVal      = objRSPIDCnt.Fields("nextval").Value
		image         = "http://sports.cbsimg.net/images/players/unknown_player.gif"
		creationDate  = date()
	
		strSQL ="insert into tblPlayers(PID,OwnerID,playerStatus,firstName,lastName,nbateamid, pos, image, clearwaiverdate,pendingwaiver,pendingtrade,rentalPlayer,OntheBlock) values ('" &_
		wNextVal & "', '" & playerCnters & "', '" & playerStatus & "', '" & barpPlayerName(0) & "', '" & barpPlayerName(1) & "', '" & nbaTeam & "', '" & position & "','" & image & "', '" &_
		date() & "', '" & playerCnters & "', '" & playerCnters & "', '" & playerCnters & "', '" & playerCnters & "')"

		objConn.Execute strSQL			
		'sURL = "commish.asp"
		'AddLinkParameter "var_ownerid", ownerid, sURL
		'Response.Redirect sURL
 end if
		
  		
%>
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
</head>
<style>
.alert-banner {
    color: #fff;
    background-color: yellowgreen;
    border-color: #000000;
    font-size: 14px;
    text-transform: none;
    border-style: double;
    border-width: medium;
}
</style>
<body>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				 <a href="#" class="close" data-dismiss="alert">&times;</a>
				<strong><span class="glyphicon glyphicon-user"></span> ADD NEW PLAYER</strong> <br>Used to Add Missing Players to the DB who have Barp Stats.
			</div>
		</div>
	</div>
</div>
<form action="playerAdd.asp" name="playerAdd" id="Add Player" method="POST">
<input type="hidden" name="action" value="Save Form Data">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
					<table class="table table-bordered table-condensed">
						<tr>
							<div class="col-md-6 col-sm-6 col-xs-6">
								<td width="50%">        
								<select class="form-control input-sm" style="color: red;" name="nbaplayername" >
									<option value="Select Player Name" selected>Select Player Name</option>
									<%
										While Not objRSplayers.EOF
									%>
									<option value="<%=objRSPlayers("first")%>;<%=objRSPlayers("last")%>"><%=objRSplayers.Fields("first")%>&nbsp;<%=objRSPlayers("last")%>&nbsp;|&nbsp;<%=objRSPlayers("team")%></option>
									<%
										objRSplayers.MoveNext
										Wend
									%>								
								</select>
							</td>
						</div>
						<div class="col-md-6 col-sm-6 col-xs-6">
							<td>        
								<select class="form-control input-sm" style="color: red;" name="playerStatus">
									<option value="Select Player Status" selected>Select Player Status</option>
									<option value="F">Free Agent</option>
									<option value="S">Staggered</option>
									<option value="W">Waivers</option>
								</select>
							</td> 
						</div>
					</tr>
					<tr>
						<div class="col-md-6 col-sm-6 col-xs-6">
							<td>        
								<select class="form-control input-sm" style="color: red;" name="teamName">
									<option value="Select NBA Team Name" selected>Select NBA Team Name</option>
									<option value="1">Atlanta Hawks</option>
									<option value="2">Boston Celtics</option>
									<option value="3">Brooklyn Nets</option>
									<option value="4">Charlotte Hornets</option>
									<option value="5">Chicago Bulls</option>
									<option value="6">Cleveland Cavaliers</option>
									<option value="7">Dallas Mavericks</option>
									<option value="8">Denver Nuggets</option>
									<option value="9">Detroit Pistons</option>
									<option value="10">Golden State Warriors</option>
									<option value="11">Houston Rockets</option>
									<option value="12">Indiana Pacers</option>
									<option value="13">Los Angeles Clippers</option>
									<option value="14">Los Angeles Lakers</option>
									<option value="15">Memphis Grizzlies</option>
									<option value="16">Miami Heat</option>
									<option value="17">Milwaukee Bucks</option>
									<option value="18">Minnesota Timberwolves</option>
									<option value="19">New Orleans Pelicans</option>
									<option value="20">New York Knicks</option>
									<option value="21">OKC Thunder</option>
									<option value="22">Orlando Magic</option>
									<option value="23">Philadelphia 76ers</option>
									<option value="24">Phoenix Suns</option>
									<option value="25">Portland Trailblazers</option>
									<option value="26">Sacramento Kings</option>
									<option value="27">San Antonio Spurs</option>
									<option value="28">Toronto Raptors</option>
									<option value="29">Utah Jazz</option>
									<option value="30">Washington Wizards</option>
								</select>
							</td>
						</div>
						<div class="col-md-6 col-sm-6 col-xs-6">
							<td>        
								<select class="form-control input-sm" style="color: red;" name="position">
									<option value="Select Position" selected>Select Position</option>
									<option value="CEN">Center</option>
									<option value="FOR">Forward</option>
									<option value="F-C">Forward-Center</option>
									<option value="GUA">Guard</option>
									<option value="G-F">Guard-Forward</option>			 
								</select>
							</td> 
						</div>
					</tr>
				</table>

	</div>
	</div>
</div>
</br>
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12" align="right">
					<button type="submit" value="Save Record" class="btn btn-trades  btn-sm"><span class="glyphicon glyphicon-save"></span>&nbsp;Save Record</button>
					<button type="reset" value="Reset" name="Reset" class="btn btn-trades  btn-sm"><span class="glyphicon glyphicon-refresh"></span>&nbsp;Refresh</button>
	</div>
</div>
</form>
<%  
	objRSPIDCnt.Close 
	objRSPlayers.Close
	Set objRSPIDCnt = Nothing
	Set objRSPlayers= Nothing 
	ObjConn.Close
	Set objConn     = Nothing
	Session.CodePage= Session("FP_OldCodePage")
	Session.LCID    = Session("FP_OldLCID")
%>
</body>
</html>