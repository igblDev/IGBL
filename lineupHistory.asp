<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear
	

	Dim objConn,objsTeams,objRS,objRSLineupsPics,firstName,lastName,image,nbaTeamID,pos,pStatus,clrWaiverDate,injury,lastTeamInd,otb,pendingWaiver,rentalPlayer,pendingTrade
	

	GetAnyParameter "Action", sAction	
		ownerid = session("ownerid")	
	
	if ownerid = "" then
		GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if	

	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"
		
	Set objsTeams = Server.CreateObject("ADODB.RecordSet")
	Set objRSLineupsPics = Server.CreateObject("ADODB.RecordSet")
	Set objRS = Server.CreateObject("ADODB.RecordSet")	
	
	select case sAction
	  case ""
			objsTeams.Open "SELECT * FROM tblowners where ownerID < 99 order by teamname asc", objConn
	  
		case "Retrieve Team"
		owner = Split(Request.Form("igblteam"),";")	
		
		'Response.Write " Owner ID = "&owner(0)&" <br> "
		objRS.Open 	"SELECT * FROM tbl_lineups_history where ownerID = " & owner(0) & " ORDER BY tbl_lineups_history.GameDay desc, tbl_lineups_history.TimeStamp desc" , objConn,3,3,1
		'Response.Write "Record Count = "& objRS.RecordCount &" <br> "
 
	end select	
	
 		
%>
	<!--#include virtual="Common/session.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, user-scalable=no">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2016-2017</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
td {
    vertical-align: middle;
  	text-align: center;
		font-size:10px;
}
black{
	color:black;
}
button{
    min-width: 20px;
    min-height: 20px;
}
.badgeFlags {
    display: inline-block;
    min-width: 10px;
    padding: 3px 6px;
    line-height: 1;
    color: white;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color: #111;
    border-radius: 14px;
    border: #111;
    /* border-style: double; */
    color: yellow;
    color: yellow;
}
</style>
</head>
<body>
<script language="JavaScript" type="text/javascript">
$("[name='my-checkbox']").bootstrapSwitch();
</script>
<!--#include virtual="Common/headerMain.inc"-->
<form action="lineupHistory.asp" name="playerMaint" id="Player" method="POST">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Lineup Histories</strong>
			</div>
		</div>
	</div>
</div>
<%	if sAction = "" then%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
				<div class="panel-body">
					<table class="table table-striped table-bordered table-condensed">
						<tr>
							<div class="col-md-12 col-sm-12 col-xs-12">
							<td>        
								<select class="form-control input-sm" name="igblteam">
									<option value="Select Player Name" selected>Select Team</option>
									<%
										While Not objsTeams.EOF
									%>
									<option value="<%=objsTeams("ownerID")%>;<%=objsTeams("teamName")%>"><%=objsTeams.Fields("teamName")%></option>
									<%
										objsTeams.MoveNext
										Wend
									%>								
								</select>
							</td>
							</div>
						</tr>
					</table>
				</div>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-sm-12 col-md-12">
					<button type="submit" name="Action"  value="Retrieve Team" class="btn btn-trades btn-block"><span class="glyphicon glyphicon-save"></span>&nbsp;Retrieve Team
		</div>
	</div>
	<br>
</div>
<%end if%>
<%if sAction = "Retrieve Team" then
	objsTeams.Open "SELECT * FROM tblowners where ownerID < 99 order by teamname asc", objConn
%>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
				<div class="panel-heading clearfix">
					<h3 class="panel-title">
						<button type="submit" name="Action"  value="Retrieve Team" class="btn btn-trades btn-xs"><span class="glyphicon glyphicon-save"></span>&nbsp;Retrieve Team
				</div>
				<div class="panel-body">
					<table class="table table-striped table-bordered table-condensed">
						<tr>
							<div class="col-md-12 col-sm-12 col-xs-12">
							<td>        
								<select class="form-control input-sm" name="igblteam">
									<option value="Select Player Name" selected>Select Team</option>
									<%
										While Not objsTeams.EOF
									%>
									<option value="<%=objsTeams("ownerID")%>;<%=objsTeams("teamName")%>"><%=objsTeams.Fields("teamName")%></option>
									<%
										objsTeams.MoveNext
										Wend
									%>								
								</select>
							</td>
							</div>
						</tr>
					</table>
				</div>
			</div>
		</div>
	</div>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div><h4><%= owner(1)%></h4></div>
		</div>
	</div>
</div>
<%
   While Not objRS.EOF
%>
<% 
	sCenter       = objRS.Fields("sCenter").Value
	sForward      = objRS.Fields("sforward").Value
	sForward2     = objRS.Fields("sforward2").Value
	sGuard        = objRS.Fields("sguard").Value
	sGuard2       = objRS.Fields("sguard2").Value
	
	sCenterTime   = objRS.Fields("sCenterTip").Value
	sForwardTime  = objRS.Fields("sForwardTip").Value
	sForward2Time = objRS.Fields("sForwardTip2").Value
	sGuardTime    = objRS.Fields("sGuardTip").Value
	sGuard2Time   = objRS.Fields("sGuardTip2").Value
	bPenalty      = objRS.Fields("penalty").Value
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sCenter  & " = PID " , objConn,3,3,1
	cFirstName    = left(objRSLineupsPics.Fields("firstName").Value,1)
	cLastName     = objRSLineupsPics.Fields("lastName").Value
	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward  & " = PID " , objConn,3,3,1
	f1FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	f1LastName    = left(objRSLineupsPics.Fields("lastName").Value,8)
	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sForward2  & " = PID " , objConn,3,3,1
	f2FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	f2LastName    = left(objRSLineupsPics.Fields("lastName").Value,8)
	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard  & " = PID " , objConn,3,3,1
	g1FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	g1LastName    = left(objRSLineupsPics.Fields("lastName").Value,8)
	objRSLineupsPics.Close
	
	objRSLineupsPics.Open "Select image, firstName, lastName from tblplayers where  " & sGuard2 & " = PID " , objConn,3,3,1
	g2FirstName   = left(objRSLineupsPics.Fields("firstName").Value,1)
	g2LastName    = left(objRSLineupsPics.Fields("lastName").Value,8)
	objRSLineupsPics.Close 

%>
<div class="container">
  <div class="row">
    <div class="col-md-12 col-sm-12 col-xs-12">
      <div class="panel panel-override">
        <table class="table table-bordered table-striped table-condensed">
          <tr>
		    <% if bPenalty = true Then %>
                <th align="center" colspan="5" style="background-color:black;color:white">*System*&nbsp;<%= objRS.Fields("timestamp").Value %> | GM Day:&nbsp;<%= objRS.Fields("gameday").Value %></th>
            <% else %>				
                <th align="center" colspan="5" style="background-color:black;color:white"><%= objRS.Fields("timestamp").Value %><span class="pull-right badgeFlags">Game:&nbsp;<%= objRS.Fields("gameday").Value %></span></th>			
            <% end if %>								
					</tr>
          <tr>
            <td width="20%"><black><%=cLastName%></black></td>
            <td width="20%"><black><%=f1LastName%></black></td>
            <td width="20%"><black><%=f2LastName%></black></td>
            <td width="20%"><black><%=g1LastName%></black></td>
            <td width="20%"><black><%=g2LastName%></black></td>
          </tr>
					<tr>
            <td width="20%"><%=sCenterTime%></td>
            <td width="20%"><%=sForwardTime%></td>
            <td width="20%"><%=sForward2Time%></td>
            <td width="20%"><%=sGuardTime%></td>
            <td width="20%"><%=sGuard2Time%></td>
          </tr>
        </table>
      </div>
    </div>
  </div>
</div>
<%
   objRS.MoveNext
   Wend
%>
<%end if%>

</form>
<%  
	objsTeams.Close
	objRSLineupsPics.Close
	objRS.Close
  Set objsTeams = Nothing
  ObjConn.Close
  Set objConn = Nothing
  Session.CodePage = Session("FP_OldCodePage")
  Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>