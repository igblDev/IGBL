<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

	
	Dim objRS,objConn, sAction, action, objRSAll, ownerid,comparePID,objrsNext5,objsName
	
	ownerid = session("ownerid")
  if ownerid = "" then
    	GetAnyParameter "var_ownerid", ownerid
	end if
	
	if ownerid = "" then
		Response.Redirect("timeout.asp")
	end if	
	
	GetAnyParameter "var_comparePID", comparePID
	
	Set objConn  = Server.CreateObject("ADODB.Connection")
	
	pid = Split((comparePID),",")		
	playerCnt = UBound(pid) + 1

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"


	Set objrsNext5 = Server.CreateObject("ADODB.RecordSet")
	GetAnyParameter "action", sAction


%>
	<!--#include virtual="Common/session.inc"-->
<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="">
<meta name="author" content="">
<title>IGBL 2017-2018</title>
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<link href="css/tabs.css" rel="stylesheet">
<style>
tr {
    vertical-align: middle;
		text-align: center;
		background-color:white;
}
th {
    vertical-align: middle;
		text-align: center;
}
red {
	color: red;
}
.badge1 {
    display:inline-block;
    min-width:10px;
    padding:3px 5px;
    font-weight:500;
    line-height:1;
    color: #fff;
    text-align:center;
    white-space:nowrap;
    vertical-align:baseline;
    background-color:#a94442;
    border-radius:14px;
}
.badge2 {
    display:inline-block;
		
		
    min-width:10px;
    padding:3px 5px;
    font-weight:500;
    line-height:1;
    color: #fff;
    text-align:center;
    white-space:nowrap;
    vertical-align:baseline;
    background-color:#8a6d3b;
    border-radius:14px;
}
.badge3 {
    display:inline-block;
    min-width:10px;
    padding:3px 5px;
    font-weight:500;
    line-height:1;
    color: #fff;
    text-align:center;
    white-space:nowrap;
    vertical-align:baseline;
    background-color:#673AB7;
    border-radius:14px;
}
.badge4 {
    display:inline-block;
    min-width:10px;
    padding:3px 5px;
    font-weight:500;
    line-height:1;
    color: #fff;
    text-align:center;
    white-space:nowrap;
    vertical-align:baseline;
    background-color:#31708f;
    border-radius:14px;
}
.panel-override {
  background-color: #ddd;
  border-color: #354478;
	border-width: 1px;
	border-radius: 0;
}
.nav-tabs {
    border-bottom: 2px solid black;
}
</style>
<script>
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip(); 
});
</script>
</head>
<body>
<div class="container">
	<div class="row">		
    <div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>LAST 5 PLAYER COMPARISON</strong>
			</div>
		</div>
	</div>
</div>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#barps">Stats</a></li>
				<li><a data-toggle="tab" href="#schedule">Schedule</a></li>
			</ul>
		</div>
	</div>
</br>	
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<div id="barps" class="tab-pane fade in active">
					<div class="panel panel-override">
							<table class="table table-custom-black table-responsive table-condensed table-bordered">
						
						<%
							Set objRSAll  = Server.CreateObject("ADODB.RecordSet")	
							Set objsName  = Server.CreateObject("ADODB.RecordSet")	
							Set objsBarps = Server.CreateObject("ADODB.RecordSet")								
							<!--FIRST PLAYER SELECTED-->
							objsName.Open "Select firstName,lastName,POS,PID,image  from tblPlayers where PID = "&pid(0)&" ", objConn,3,3,1
							
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
							objsName.Open "Select firstName,lastName,POS,PID,image  from tblPlayers where PID = "&pid(1)&" ", objConn,3,3,1
									
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
						<!--EOF FIRST PLAYER SELECTED-->								

						<!--THIRD PLAYER SELECTED-->
							if playerCnt >= 3 then
							objsName.Open "Select firstName,lastName,POS,PID,image  from tblPlayers where PID = "&pid(2)&" ", objConn,3,3,1
									
							wFirstName3   = objsName.Fields("firstName").Value
							wLastName3    = objsName.Fields("lastName").Value
							intFName3     = left(wFirstName3,2)
							intLName3			= left(wLastName3,2)
							image3        = objsName.Fields("image").Value
							wPos3         = objsName.Fields("POS").Value
							
							objsBarps.Open "SELECT barps,usage FROM tbl_barps t WHERE t.first = '" & wFirstName3 & "'  and t.last ='"&wLastName3 & "' ", objConn,1,1					

							objRSAll.Open "SELECT avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
														"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP " &_
														"FROM tblLast5 t " & _
														"WHERE t.first = '" & wFirstName3 & "'  and t.last ='"&wLastName3 & "' ", objConn,1,1
					
							avgMP3    = objRSAll.Fields("avgMP").Value
							avgBlks3  = objRSAll.Fields("avgBlks").Value
							avgAst3   = objRSAll.Fields("avgAst").Value
							avgReb3   = objRSAll.Fields("avgReb").Value
							avgPts3   = objRSAll.Fields("avgPts").Value
							avgStl3   = objRSAll.Fields("avgStl").Value
							avg3pt3   = objRSAll.Fields("avg3pt").Value
							avgTo3    = objRSAll.Fields("avgTo").Value
							avgBarps3 = objRSAll.Fields("avgBarps").Value	
							barps3    = objsBarps.Fields("Barps").Value 	
							usage3    = objsBarps.Fields("usage").Value 	
							
							objRSAll.close	
							objsName.close	
							objsBarps.close	
						else
							avgMP3    = 0
							avgBlks3  = 0
							avgAst3   = 0
							avgReb3   = 0
							avgPts3   = 0
							avgStl3   = 0
							avg3pt3   = 0
							avgTo3    = 99
							avgBarps3 = 0
							barps3    = 0 
							usage3    = 0 
						end if		
						<!--EOF THIRD PLAYER SELECTED-->		
						<!--FOURTH PLAYER SELECTED-->
							if playerCnt >= 4 then
							objsName.Open "Select firstName,lastName,POS,PID,image  from tblPlayers where PID = "&pid(3)&" ", objConn,3,3,1
									
							wFirstName4   = objsName.Fields("firstName").Value
							wLastName4    = objsName.Fields("lastName").Value
							intFName4     = left(wFirstName4,2)
							intLName4			= left(wLastName4,2)
							image4        = objsName.Fields("image").Value
							wPos4         = objsName.Fields("POS").Value
							objsBarps.Open "SELECT barps,usage FROM tbl_barps t WHERE t.first = '" & wFirstName4 & "'  and t.last ='"&wLastName4 & "' ", objConn,1,1					

							objRSAll.Open "SELECT avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
														"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP " &_
														"FROM tblLast5 t " & _
														"WHERE t.first = '" & wFirstName4 & "'  and t.last ='"&wLastName4 & "' ", objConn,1,1
					
							avgMP4    = objRSAll.Fields("avgMP").Value
							avgBlks4  = objRSAll.Fields("avgBlks").Value
							avgAst4   = objRSAll.Fields("avgAst").Value
							avgReb4   = objRSAll.Fields("avgReb").Value
							avgPts4   = objRSAll.Fields("avgPts").Value
							avgStl4   = objRSAll.Fields("avgStl").Value
							avg3pt4   = objRSAll.Fields("avg3pt").Value
							avgTo4    = objRSAll.Fields("avgTo").Value
							avgBarps4 = objRSAll.Fields("avgBarps").Value	
							barps4    = objsBarps.Fields("Barps").Value 
							usage4    = objsBarps.Fields("usage").Value 
							
							objRSAll.close	
							objsName.close	
							objsBarps.close	
						else 
							avgMP4    = 0
							avgBlks4  = 0
							avgAst4   = 0
							avgReb4   = 0
							avgPts4   = 0
							avgStl4   = 0
							avg3pt4   = 0
							avgTo4    = 99
							avgBarps4 = 0
							barps4    = 0
							uage4     = 0 		

						
						end if		
						<!--EOF FOURTH PLAYER SELECTED-->	
						%>

	
								<tr style="background-color:white;vertical-align:middle;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;vertical-align:middle;">Player</th>
								<td style="vertical-align:middle;background-color:white;color:yellowgreen;font-weight:bold;"><img class="img-responsive center-block" src="<%=image%>"></td>	
								<td style="vertical-align:middle;background-color:white;color:yellowgreen;font-weight:bold;"><img class="img-responsive center-block" src="<%=image2%>"></td>
								<% if playerCnt >= 3 then %>								
								<td style="vertical-align:middle;background-color:white;color:yellowgreen;font-weight:bold;"><img class="img-responsive center-block" src="<%=image3%>"></td>	
								<%else%>
								<td style="vertical-align:middle;background-color:white;color:yellowgreen;font-weight:bold;"><img class="img-responsive center-block" src="http://i.cdn.turner.com/nba/nba/.element/img/2.0/sect/statscube/players/large/default_nba_headshot_v2.png"</td>
								<%end if%>
								<% if playerCnt >= 4 then %>	
								<td style="vertical-align:middle;background-color:white;color:yellowgreen;font-weight:bold;"><img class="img-responsive center-block" src="<%=image4%>"></td>	
								<%else%>
								<td style="vertical-align:middle;background-color:white;color:yellowgreen;font-weight:bold;"><img class="img-responsive center-block" src="http://i.cdn.turner.com/nba/nba/.element/img/2.0/sect/statscube/players/large/default_nba_headshot_v2.png"></td>
								<%end if%>
							</tr>
							<tr class="big" style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>POS</strong></th>
								<th class="big" style="width:20%;background-color: yellowgreen;font-weight: bold;"><%=wPos%></th>	
								<th class="big" style="width:20%;background-color: yellowgreen;font-weight: bold;"><%=wPos2%></th>
								<% if playerCnt >= 3 then %>	
								<th class="big"  style="width:20%;background-color: yellowgreen;font-weight: bold;"><%=wPos3%></th>
								<%else%>
								<th class="big"  style="width:20%;background-color: yellowgreen;font-weight: bold;">N/A</th>
								<%end if %>	
								<% if playerCnt >= 4 then %>									
									<th class="big"  style="width:20%;background-color: yellowgreen;font-weight: bold;"><%=wPos4%></th>				
								<%else%>
									<th class="big"  style="width:20%;background-color: yellowgreen;font-weight: bold;">N/A</th>
								<%end if%>	
							</tr>								
							<tr class="big" style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>MPG</strong></th>
								<% if avgMP > avgMP2 and  avgMP > avgMP3 and  avgMP > avgMP4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;width:20%;"><%=round(avgMP,2)%></td>	
								<%else%>
									<td class="big"  style="width:20%"><%=round(avgMP,2)%></td>								
								<%end if %>
								
								<% if avgMP2 > avgMP and  avgMP2 > avgMP3 and  avgMP2 > avgMP4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;width:20%;"><%=round(avgMP2,2)%></td>	
								<%else%>
									<td class="big"  style="width:20%"><%=round(avgMP2,2)%></td>								
								<%end if %>

								<% if avgMP3 > avgMP and  avgMP3 > avgMP2 and  avgMP3 > avgMP4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;width:20%;"><%=round(avgMP3,2)%></td>	
								<%else%>
									<td class="big"  style="width:20%"><%=round(avgMP3,2)%></td>								
								<%end if %>
								
								<% if avgMP4 > avgMP and  avgMP4 > avgMP2 and  avgMP4 > avgMP3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;width:20%;"><%=round(avgMP4,2)%></td>	
								<%else%>
									<td class="big"  style="width:20%"><%=round(avgMP4,2)%></td>								
								<%end if %>
							</tr>						
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>BPG</strong></th>	
								<% if avgBlks > avgBlks2 and  avgBlks > avgBlks3 and  avgBlks > avgBlks4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;width:20%;"><%=round(avgBlks,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBlks,2)%></td>								
								<%end if %>
								<% if avgBlks2 > avgBlks and  avgBlks2 > avgBlks3 and  avgBlks2 > avgBlks4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;width:20%;"><%=round(avgBlks2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBlks2,2)%></td>								
								<%end if %>
								<% if avgBlks3 > avgBlks and  avgBlks3 > avgBlks2 and  avgBlks3 > avgBlks4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avgBlks3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBlks3,2)%></td>								
								<%end if %>								
								
								<% if avgBlks4 > avgBlks and  avgBlks4 > avgBlks2 and  avgBlks4 > avgBlks3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(avgBlks4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBlks4,2)%></td>								
								<%end if %>	
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>APG</strong></th>	
								<% if avgAst > avgAst2 and  avgAst > avgAst3 and  avgAst > avgAst4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(avgAst,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgAst,2)%></td>								
								<%end if %>
								
								<% if avgAst2 > avgAst and  avgAst2 > avgAst3 and  avgAst2 > avgAst4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(avgAst2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgAst2,2)%></td>								
								<%end if %>
								
								<% if avgAst3 > avgAst and  avgAst3 > avgAst2 and  avgAst3 > avgAst4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avgAst3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgAst3,2)%></td>								
								<%end if %>								
								
								<% if avgAst4 > avgAst and  avgAst4 > avgAst2 and  avgAst4 > avgAst3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%= round(avgAst4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgAst4,2)%></td>								
								<%end if %>
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>RPG</strong></th>	
								<% if avgReb > avgReb2 and  avgReb > avgReb3 and  avgReb > avgReb4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(avgReb,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgReb,2)%></td>								
								<%end if %>
								
								<% if avgReb2 > avgReb and  avgReb2 > avgReb3 and  avgReb2 > avgReb4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(avgReb2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgReb2,2)%></td>								
								<%end if %>
								
								<% if avgReb3 > avgReb and  avgReb3 > avgReb2 and  avgReb3 > avgReb4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avgReb3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgReb3,2)%></td>								
								<%end if %>								
								
								<% if avgReb4 > avgReb and  avgReb4 > avgReb2 and  avgReb4 > avgReb3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(avgReb4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgReb4,2)%></td>								
								<%end if %>
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>PPG</strong></th>	
								<% if avgPts > avgPts2 and  avgPts > avgPts3 and  avgPts > avgPts4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(avgPts,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgPts,2)%></td>								
								<%end if %>
								<% if avgPts2 > avgPts and  avgPts2 > avgPts3 and  avgPts2 > avgPts4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(avgPts2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgPts2,2)%></td>								
								<%end if %>
								<% if avgPts3 > avgPts and  avgPts3 > avgPts2 and  avgPts3 > avgPts4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avgPts3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgPts3,2)%></td>								
								<%end if %>								
								
								<% if avgPts4 > avgPts and  avgPts4 > avgPts2 and  avgPts4 > avgPts3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(avgPts4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgPts4,2)%></td>								
								<%end if %>
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>SPG</strong></th>	
								<% if avgStl > avgStl2 and  avgStl > avgStl3 and  avgStl > avgStl4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(avgStl,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgStl,2)%></td>								
								<%end if %>
								<% if avgStl2 > avgStl and  avgStl2 > avgStl3 and  avgStl2 > avgStl4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(avgStl2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgStl2,2)%></td>								
								<%end if %>
								<% if avgStl3 > avgStl and  avgStl3 > avgStl2 and  avgStl3 > avgStl4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avgStl3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgStl3,2)%></td>								
								<%end if %>								

								<% if avgStl4 > avgStl and  avgStl4 > avgStl2 and  avgStl4 > avgStl3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(avgStl4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgStl4,2)%></td>								
								<%end if %>
  						</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>3PG</strong></th>	
								<% if avg3pt > avg3pt2 and  avg3pt > avg3pt3 and  avg3pt > avg3pt4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(avg3pt,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avg3pt,2)%></td>								
								<%end if %>
								<% if avg3pt2 > avg3pt and  avg3pt2 > avg3pt3 and  avg3pt2 > avg3pt4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(avg3pt2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avg3pt2,2)%></td>								
								<%end if %>
								<% if avg3pt3 > avg3pt and  avg3pt3 > avg3pt2 and  avg3pt3 > avg3pt4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avg3pt3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avg3pt3,2)%></td>								
								<%end if %>								
								
								<% if avg3pt4 > avg3pt and  avg3pt4 > avg3pt2 and  avg3pt4 > avg3pt3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(avg3pt4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avg3pt4,2)%></td>								
								<%end if %>	
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>TPG</strong></th>	
								<% if avgTo < avgTo2 and  avgTo < avgTo3 and  avgTo < avgTo4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(avgTo,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgTo,2)%></td>								
								<%end if %>
								<% if avgTo2 < avgTo and  avgTo2 < avgTo3 and  avgTo2 < avgTo4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(avgTo2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgTo2,2)%></td>								
								<%end if %>
								<% if avgTo3 < avgTo and  avgTo3 < avgTo2 and  avgTo3 < avgTo4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avgTo3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgTo3,2)%></td>								
								<%end if %>								
								
								<% if avgTo4 < avgTo and  avgTo4 < avgTo2 and  avgTo4 < avgTo3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(avgTo4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgTo4,2)%></td>								
								<%end if %>	
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>Last 5</strong></th>	
								<% if avgBarps > avgBarps2 and  avgBarps > avgBarps3 and  avgBarps > avgBarps4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(avgBarps,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBarps,2)%></td>								
								<%end if %>
								<% if avgBarps2 > avgBarps and  avgBarps2 > avgBarps3 and  avgBarps2 > avgBarps4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(avgBarps2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBarps2,2)%></td>								
								<%end if %>
								<% if avgBarps3 > avgBarps and  avgBarps3 > avgBarps2 and  avgBarps3 > avgBarps4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(avgBarps3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBarps3,2)%></td>								
								<%end if %>								
								
								<% if avgBarps4 > avgBarps and  avgBarps4 > avgBarps2 and  avgBarps4 > avgBarps3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(avgBarps4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(avgBarps4,2)%></td>								
								<%end if %>								
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>Usage</strong></th>	
								<% if usage > usage2 and  usage > usage3 and  usage > usage4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(usage,2)%></td>	
								<%else%>
									<td class="big" ><%=round(usage,2)%></td>								
								<%end if %>
								<% if usage2 > usage and  usage2 > usage3 and  usage2 > usage4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(usage2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(usage2,2)%></td>								
								<%end if %>
								<% if usage3 > usage and  usage3 > usage2 and  usage3 > usage4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(usage3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(usage3,2)%></td>								
								<%end if %>								
								
								<% if usage4 > usage and  usage4 > usage2 and  usage4 > usage3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(usage4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(usage4,2)%></td>								
								<%end if %>								
							</tr>	
							<tr style="background-color:white;">
								<th class="big" style="color:black;font-weight:bold;background-color:silver;"><strong>Barps</strong></th>	
								<% if barps > barps2 and  barps > barps3 and  barps > barps4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#a94442;"><%=round(barps,2)%></td>	
								<%else%>
									<td class="big" ><%=round(barps,2)%></td>								
								<%end if %>
								<% if barps2 > barps and  barps2 > barps3 and  barps2 > barps4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#8a6d3b;"><%=round(barps2,2)%></td>	
								<%else%>
									<td class="big" ><%=round(barps2,2)%></td>								
								<%end if %>
								<% if barps3 > barps and  barps3 > barps2 and  barps3 > barps4 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#3c763d;"><%=round(barps3,2)%></td>	
								<%else%>
									<td class="big" ><%=round(barps3,2)%></td>								
								<%end if %>								
								
								<% if barps4 > barps and  barps4 > barps2 and  barps4 > barps3 then %>	
									<td class="big"  style="color:white;font-weight:bold;background-color:#31708f;"><%=round(barps4,2)%></td>	
								<%else%>
									<td class="big" ><%=round(barps4,2)%></td>								
								<%end if %>	
							</tr>	
						</tbody>
					</table>
					</div>
				</div>
				<div id="schedule" class="tab-pane fade">
				<%
					objrsNext5.Open "SELECT * FROM qryAllPlayerGameDays where PID in("&pid(0)&") and gameday >= Date() order by pid,gameday ", objConn,1,1
					skedcnt = 1
				%>
					<div class="panel panel-override">
						<table class="table table-responsive table-custom-black table-bordered table-condensed" >
							<thead>
							 <tr style="color:#354478">
								<th style="width:20%;"><strong>GM 1</strong></th>
								<th style="width:20%;"><strong>GM 2</strong></th>
								<th style="width:20%;"><strong>GM 3</strong></th>
								<th style="width:20%;"><strong>GM 4</strong></th>
								<th style="width:20%;"><strong>GM 5</strong></th>
								</tr>
							</thead>
							<tbody>
								<%
								While Not objRSNext5.EOF and skedcnt <5
								%>									
									<tr class="danger">
										<td style="vertical-align:middle;text-align:center;color:#a94442;font-weight:bold;" colspan="5"><%=objRSNext5.Fields("firstName").Value%>&nbsp;<%=objRSNext5.Fields("lastName").Value%>&nbsp;-&nbsp;<%=objRSNext5.Fields("teamshortName").Value%></td>
									</tr>
									<tr style="background-color:white;">
									<%While skedcnt <=5 %>
										<td style="vertical-align:middle;text-align:center;width:5%;color:#616161;"><%=objRSNext5.Fields("opponent").Value%></br><%=objRSNext5.Fields("gameDay").Value%></td>
									<%
										objRSNext5.MoveNext
										skedcnt = skedcnt + 1
										Wend
									%>
									</tr>
								<%
									objrsNext5.MoveNext
									'skedcnt = 1
										Wend
								%>

								<%
								objrsNext5.Close
								objrsNext5.Open "SELECT * FROM qryAllPlayerGameDays where PID in("&pid(1)&") and gameday >= Date() order by pid,gameday ", objConn,1,1
								skedcnt = 1
								While Not objRSNext5.EOF and skedcnt <5
								%>
									<tr class="warning">
										<td style="vertical-align:middle;text-align:center;color:#8a6d3b;font-weight:bold;" colspan="5"><%=objRSNext5.Fields("firstName").Value%>&nbsp;<%=objRSNext5.Fields("lastName").Value%>&nbsp;-&nbsp;<%=objRSNext5.Fields("teamshortName").Value%></td>
									</tr>
									<tr style="background-color:white;">
									<%While skedcnt <=5 %>
										<td style="vertical-align:middle;text-align:center;width:5%;color:#616161;"><%=objRSNext5.Fields("opponent").Value%></br><%=objRSNext5.Fields("gameDay").Value%></td>
									<%
										objRSNext5.MoveNext
										skedcnt = skedcnt + 1
										Wend
									%>
									</tr>
								<%
									objrsNext5.MoveNext
										Wend
								%>

								<%if playerCnt >= 3 then %>
									<%
									objrsNext5.Close
									objrsNext5.Open "SELECT * FROM qryAllPlayerGameDays where PID in("&pid(2)&") and gameday >= Date() order by pid,gameday ", objConn,1,1
									skedcnt = 1
									While Not objRSNext5.EOF and skedcnt <5
									'While skedcnt <5 and Not objRSNext5.eof
									%>
										<tr class="success">
											<td style="vertical-align:middle;text-align:center;color:#3c763d;font-weight:bold;" colspan="5"><%=objRSNext5.Fields("firstName").Value%>&nbsp;<%=objRSNext5.Fields("lastName").Value%>&nbsp;-&nbsp;<%=objRSNext5.Fields("teamshortName").Value%></td>
										</tr>
										<tr style="background-color:white;">
										<%While skedcnt <=5 %>
											<td style="vertical-align:middle;text-align:center;width:5%;color:#616161;"><%=objRSNext5.Fields("opponent").Value%></br><%=objRSNext5.Fields("gameDay").Value%></td>
										<%
											objRSNext5.MoveNext
											skedcnt = skedcnt + 1
											Wend
										%>
										</tr>
									<%
										objrsNext5.MoveNext
											Wend
									%>
								<%end if %>	

								<%if playerCnt >= 4 then %>
									<%
									objrsNext5.Close
									objrsNext5.Open "SELECT * FROM qryAllPlayerGameDays where PID in("&pid(3)&") and gameday >= Date() order by pid,gameday ", objConn,1,1
									skedcnt = 1
									While Not objRSNext5.EOF and skedcnt <5
									%>
										<tr class="info">
											<td style="vertical-align:middle;text-align:center;color:#31708f;font-weight:bold;" colspan="5"><%=objRSNext5.Fields("firstName").Value%>&nbsp;<%=objRSNext5.Fields("lastName").Value%>&nbsp;-&nbsp;<%=objRSNext5.Fields("teamshortName").Value%></td>
										</tr>
										<tr style="background-color:white;">
										<%While skedcnt <=5 %>
											<td style="vertical-align:middle;text-align:center;width:5%;color:#616161;"><%=objRSNext5.Fields("opponent").Value%></br><%=objRSNext5.Fields("gameDay").Value%></td>
										<%
											objRSNext5.MoveNext
											skedcnt = skedcnt + 1
											Wend
										%>
										</tr>
									<%
										objrsNext5.MoveNext
											Wend
									%>
								<%end if %>	
								</tbody>
						</table>
						<%
							Set objRSgames = Server.CreateObject("ADODB.RecordSet")		
							objRSgames.Open "qryGameDeadLines", objConn	
							skedcnt = 1
						%>
						<table class="table table-responsive table-custom-black table-bordered table-condensed">
							<tr style="color:white;background-color:black;">
								<td colspan="5"><strong>Next 5 IGBL GAME DATES</strong></td>							
							</tr>
						 <tr >
								<th style="width:20%;">GM 1</th>
								<th style="width:20%;">GM 2</th>
								<th style="width:20%;">GM 3</th>
								<th style="width:20%;">GM 4</th>
								<th style="width:20%;">GM 5</th>
							</tr>
							<tr class="default">
							<% While skedcnt < 6 %>
								<td style="width:20%;color:#616161"><%=objRSgames("gameDay")%></td>
							<% 
							objRSgames.MoveNext
							skedcnt = skedcnt + 1
							Wend 
							%>
							</tr>
						</table>
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<%
  objConn.Close
	objRSgames.close
	objrsNext5.close
	objRSAll.close
	objsName.close	
	objsBarps.close	
  Set objConn = Nothing
%>
</body>
</html>