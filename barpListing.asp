<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

	Dim objRS,objConn, sAction, action, objRSCen,objRSFC,objRSGF,objRSFor,objRSGua,ownerid,objRSAll,objRSPlayerInfo

	ownerid = session("ownerid")	
	
	if ownerid = "" then
    	GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	else
		GetAnyParameter "Action", sAction	
	end if	
	
	select case sAction
	
	case "Compare"
		compareCBoxCnt = Request.Form("compPID").count
		
		if compareCBoxCnt > 1 then 
			wcomparePID = Split(Request.Form("compPID"), ";")
			
			For I = 1 To compareCBoxCnt
				wcomparePID = Split(Request.Form("compPID")(I), ";")
				if I = compareCBoxCnt then
					comparePID = comparePID &wcomparePID(0)
				else
					comparePID = comparePID &wcomparePID(0)&","
				end if
				Response.Write "Compare Check Boxes Values	 = "&comparePID&".<br>"
			NEXT
			sURL = "compareTabs.asp"
			AddLinkParameter "var_ownerid", ownerid, sURL
			AddLinkParameter "var_comparePID", comparePID, sURL
			Response.Redirect sURL
		else 
			errorCode = "No Buttons Checked"
		end if
	 
	end Select
		
	dim loopcnt
	Set objConn  = Server.CreateObject("ADODB.Connection")


	objConn.Open Application("lineupstest_ConnectionString")
  objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"

	
	Set objRSCen        = Server.CreateObject("ADODB.RecordSet")
	Set objRSFor        = Server.CreateObject("ADODB.RecordSet")
	Set objRSGua        = Server.CreateObject("ADODB.RecordSet")
	
	GetAnyParameter "action", sAction
	
	objRSCen.Open  "Select * FROM qry_PlayerAll WHERE firstName <> 'No' AND POS = 'CEN' Or POS = 'F-C' Order By barps desc" ,objConn,3,3,1
	objRSFor.Open  "Select * FROM qry_PlayerAll WHERE firstName <> 'No' AND POS = 'FOR' Order By barps desc" ,objConn,3,3,1
	objRSGua.Open  "Select * FROM qry_PlayerAll WHERE firstName <> 'No' AND POS = 'GUA' Or POS = 'G-F' Order By barps desc" ,objConn,3,3,1
	
	loopcntGua = 0 
	loopCntCen = 0
	loopCntFor = 0 

%>
	<!--#include virtual="Common/session.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Custom styles for this template -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
small {
	color:black;
}

white {
	color: white;
}
red {
	color: red;
}
gold {
	color: gold;
	font-weight: bold;
  text-transform: uppercase;
}
yellow {
	color: yellow;
}
black {
	color:black;
}

.btn-success {
    color: #FFFFFF !important;
    background-color: #468847 !important;
    border-color: #354478 !important;
}
td {
    vertical-align: middle;
}
.badge {
    display: inline-block;
    min-width: 10px;
    padding: 3px 7px;
    line-height: 1;
    color:white;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color:#354478;
    border-radius: 14px;
}
.badgeGames {
    display: inline-block;
    min-width: 10px;
    padding: 3px 7px;
    line-height: 1;
    color:yellow;
    text-align: center;
    white-space: nowrap;
    vertical-align: baseline;
    background-color:black;
    border-radius: 14px;
}
</style>
</head>
<body>
<script language="JavaScript" type="text/javascript">
function functionCompare(theForm) {
	var totalPlayers = 0;

	for (var i = 0; i < theForm.compPID.length; i++) {
    if (theForm.compPID[i].checked) {
      totalPlayers += 1;
    }
  }

	if (totalPlayers <= 1) {
		alert("Select minimum of 2 players to compare!"); 
    return false;
  }	
	
	if (totalPlayers >4) {
		alert("Select maximum of 4 players to compare!"); 
    return false;
  }	
	
return (true);
}
</script>
<!--#include virtual="Common/headermain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fal fa-balance-scale"></i>&nbsp;Player Comparasion</strong>
			</div>
		</div>
	</div>
</div>
<form action="barpListing.asp" method="POST" onSubmit="return functionCompare(this)" name="frmBarps" language="JavaScript">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
<div class="container">
	<div class="row">
		<div class="col-md-6 col-sm-6 col-xs-6">
			<button type="submit" id="idRetrieveLineup" value="Compare" name="Action" class="btn btn-default btn-block btn-sm"><i class="fa fa-balance-scale" aria-hidden="true"></i>&nbsp;Compare Players</button>
		</div>
		<div class="col-md-6 col-sm-6 col-xs-6">
			<button type="reset"  value="Reset" name="Reset" class="btn btn-default btn-block btn-sm"><i class="fa fa-repeat" aria-hidden="true"></i>&nbsp;Reset</button>
		</div>
	</div>
</div>
</br>
<div class="container">
	<div class="panel panel-override">
		<div class="panel-body">
			<table class="table table-striped table-bordered table-condensed">
				<tr bgcolor="#FFFFFF">
					<td>
						<table class="table table-striped table-responsive table-bordered table-custom-black table-condensed" style="display: block; height: 675px;overflow-y: scroll;">
							<th colspan="2">CEN | <span style="color:blue;">F-C</span></th>
						<%
						While loopCntCen <75 
						%>
							<tr>
								<td style="vertical-align:middle;text-align:center;width:5%;color:blue"><input type="checkbox" class="checkbox" name="compPID" value="<%=objRSCen.Fields("pid").Value & ";"%>"></td>
								<% if objRSCen.Fields("POS").Value = "F-C" then %>
									<td style="vertical-align:middle;color:blue;"><%=left(objRSCen.Fields("lastName").Value,10)%></br><span style="color:#468847;"><%=objRSCen.Fields("teamshortname").Value%></span>&nbsp;<span style="color:black;"><%=round(objRSCen.Fields("barps").Value,2)%></span></td>
								<% else %>
									<td style="vertical-align:middle;"><%=left(objRSCen.Fields("lastName").Value,10)%></br><span style="color:#468847;"><%=objRSCen.Fields("teamshortname").Value%></span>&nbsp;<span style="color:black;"><%=round(objRSCen.Fields("barps").Value,2)%></span></td>
								<% end if%>
							</tr>
						<%
						loopCntCen = loopCntCen + 1
						objRSCen.MoveNext
						Wend
						%>
						</table>
					</td>
					<td>
						<table class="table table-striped table-responsive table-bordered table-custom-black table-condensed" style="display: block; height: 675px;overflow-y: scroll;">
							<th colspan="2">Forwards</th>
						<%
						While loopCntFor < 75 
						%>
							<tr> 
								<td style="vertical-align:middle;text-align:center;width:5%;"><input type="checkbox" class="checkbox" name="compPID" value="<%=objRSFOR.Fields("pid").Value & ";"%>"></td>
								<td><%=left(objRSFOR.Fields("lastName").Value,10)%></br><span style="color:#468847;"><%=objRSFOR.Fields("teamshortname").Value%></span>&nbsp;<span style="color:black;"><%=round(objRSFOR.Fields("barps").Value,2)%></span></td>
							</tr>
						<%
						loopCntFor = loopCntFor + 1
						objRSFor.MoveNext
						Wend
						%>
						</table>
					</td>								
					<td>
						<table class="table table-striped table-responsive table-bordered table-custom-black table-condensed" style="display: block; height: 675px;overflow-y: scroll;">
							<th colspan="2">GUA | <span style="color:blue;">G-F</span></th></th>
						<%
						While loopCntGua < 75 
						%>
							<tr>
								<td style="vertical-align:middle;text-align:center;width:5%;"><input type="checkbox" class="checkbox" name="compPID" value="<%=objRSGua.Fields("pid").Value & ";"%>"></td>
								<% if objRSGua.Fields("POS").Value = "G-F" then %>
									<td style="vertical-align:middle;color:blue;"><%=left(objRSGua.Fields("lastName").Value,10)%></br><span style="color:#468847;"><%=objRSGua.Fields("teamshortname").Value%></span>&nbsp;<span style="color:black;"><%=round(objRSGua.Fields("barps").Value,2)%></span></td>
								<% else %>
									<td style="vertical-align:middle;"><%=left(objRSGua.Fields("lastName").Value,10)%></br><span style="color:#468847;"><%=objRSGua.Fields("teamshortname").Value%></span>&nbsp;<span style="color:black;"><%=round(objRSGua.Fields("barps").Value,2)%></span></td>
								<% end if%>							
							</tr>
						<%
						loopCntGua = loopCntGua + 1
						objRSGua.MoveNext
						Wend
						%>
						</table>
					</td>
				</tr>
			</table>
		</div>
	</div>
</div>
</form>
<%
objRSCen.close
objRSFC.close
objRSGF.close
objRSGua.close
objRSGua.close
objconn.close
Set objRS= Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>