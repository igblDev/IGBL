</br>
<%
	Set objRBox       = Server.CreateObject("ADODB.RecordSet")
	Set objrsFinalBox = Server.CreateObject("ADODB.RecordSet")
	Set objrsTeamRecH = Server.CreateObject("ADODB.RecordSet")
	Set objrsTeamRecA = Server.CreateObject("ADODB.RecordSet")
	objRBox.Open "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"# order by game_Number",objConn,3,3,1	
	
	objrsFinalBox.Open "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"# and FinalBox = true",objConn,3,3,1	
	finalBoxCnt = objrsFinalBox.RecordCount
	objrsFinalBox.Close
%>
		<table class="table table-custom-black table-responsive table-bordered table-condensed">
			<thead>
				<%if finalBoxCnt = 5 then %>
					<th class="big" colspan="3"><span class="green"><i class="fal fa-flag-checkered fa-lg" style="vertical-align:middle;"></i></span>&nbsp;<%=gameDate%>&nbsp;<i class="fal fa-calendar-alt"></i>&nbsp;BOX SCORE SUMMARY&nbsp;<span style="color:#468847;font-weight:bold;"><span style="vertical-align:middle;">[FINAL]</span></th>		
				<%else%>
					<th class="big" colspan="3"><%=gameDate%>&nbsp;<i class="fal fa-calendar-alt"></i>&nbsp;BOX SCORE SUMMARY <span class="red">[PARTIAL]</th>		
				<%end if%>
			</thead>
			<%
				While Not objRBox.EOF
					aOID = objRBox.Fields("awayTeamOID").Value
					hOID = objRBox.Fields("homeTeamOID").Value
					
					hBarpLDR =  get_top_barper(gameDate, hOID, hpName, hpBarps, hpid)						
					aBarpLDR =  get_top_barper(gameDate, aOID, apName, apBarps, apid)		
					objrsTeamRecA.Open   "SELECT * FROM standings where ID = "&aOID&"", objConn,3,3,1
					objrsTeamRecH.Open   "SELECT * FROM standings where ID = "&hOID&"", objConn,3,3,1
					
			%>
			<%if aOID = ownerid or hOID = ownerid then%>
				<tr class="success">	
			<%else%>
				<tr style="background-color:white;vertical-align:middle;">
			<%end if%>		
				<%if objRBox.Fields("htotBarps").value > objRBox.Fields("atotBarps").value then%>
					<td>
						<table class="table table-responsive table-bordered table-condensed">
							<tr style="text-align:left;vertical-align:middle;font-weight:700;text-transform:uppercase;color:black;">
								<td class="big" style="width:30%;text-align:left;vertical-align:middle;"><i class="fas fa-caret-circle-right red fa-lg"></i>&nbsp;<span style="text-transform:uppercase;"><%=objRBox.Fields("hometeamshort").value%></span>&nbsp;<%=objrsTeamRecH.Fields("won").Value %>-<%=objrsTeamRecH.Fields("loss").Value %></td>
								<!--<td class="big" style="vertical-align: middle;width:13%;"><%=objrsTeamRecH.Fields("won").Value %>-<%=objrsTeamRecH.Fields("loss").Value %></td>-->
								<td class="big" style="width:15%;background-color: black;color: #fff999;"><%=objRBox.Fields("htotBarps").value%></td>
								<td class="big" style="width:45%;text-align:left;text-transform:none;"><i class="fas fa-basketball-ball" style="color:darkorange;"></i>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=hpid%>"><%=left(hpName,14)%></a>&nbsp;<%=hpBarps%></td>
							</tr>
							<tr style="text-align:left;vertical-align:middle;color:#a5a6a7;font-weight:bold;text-transform:uppercase;">	
								<td class="big" style="width:30%;text-align:left;vertical-align:middle;color:#a5a6a7;font-weight:bold;"><%=objRBox.Fields("awayteamshort").value%>&nbsp;<%=objrsTeamRecA.Fields("won").Value %>-<%=objrsTeamRecA.Fields("loss").Value %></td>
								<!--<td class="big" style="vertical-align: middle;width:13%;"><%=objrsTeamRecA.Fields("won").Value %>-<%=objrsTeamRecA.Fields("loss").Value %></td>-->
								<td class="big" style="vertical-align: middle;width:15%;"><%=objRBox.Fields("atotBarps").value%></td>
								<td class="big" style="width:45%;text-align:left;text-transform:none;"><i class="fas fa-basketball-ball" style="color:darkorange;"></i>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=apid%>"><%=left(apName,14)%></a>&nbsp;<%=apBarps%></span></td>
							</tr>
						</table>					
					</td>
				<%elseif objRBox.Fields("atotBarps").value > objRBox.Fields("htotBarps").value then%>
					<td>
						<table class="table table-responsive table-bordered table-condensed">
							<tr style="text-align:left;vertical-align:middle;color:#a5a6a7;font-weight:bold;text-transform:uppercase;">
								<td class="big" style="width:30%;text-align:left;vertical-align:middle;color:#a5a6a7;font-weight:bold;"><%=objRBox.Fields("hometeamshort").value%>&nbsp;<%=objrsTeamRecH.Fields("won").Value %>-<%=objrsTeamRecH.Fields("loss").Value %></td>
								<!--<td class="big" style="vertical-align: middle;width:13%;"><%=objrsTeamRecH.Fields("won").Value %>-<%=objrsTeamRecH.Fields("loss").Value %></td>-->
								<td class="big" style="vertical-align: middle;width:15%;"><%=objRBox.Fields("htotBarps").value%></td>
								<td class="big" style="width:45%;text-align:left;text-transform:none;"><i class="fas fa-basketball-ball" style="color:darkorange;"></i>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=hpid%>"><%=left(hpName,14)%></a>&nbsp;<%=hpBarps%></td>
							</tr>
							<tr style="text-align:left;vertical-align:middle;font-weight:700;text-transform:uppercase;color:black;">
								<td class="big" style="width:30%;text-align:left;vertical-align:middle;"><i class="fas fa-caret-circle-right red fa-lg"></i>&nbsp;<span style="text-transform:uppercase;"><%=objRBox.Fields("awayteamshort").value%></span>&nbsp;<%=objrsTeamRecA.Fields("won").Value %>-<%=objrsTeamRecA.Fields("loss").Value %></td>
								<!--<td class="big" style="vertical-align: middle;width:13%;"><%=objrsTeamRecA.Fields("won").Value %>-<%=objrsTeamRecA.Fields("loss").Value %></td>-->
								<td class="big" style="width:15%;background-color: black;color: #fff999;"><%=objRBox.Fields("atotBarps").value%></td>
								<td class="big" style="width:45%;text-align:left;text-transform:none;"><i class="fas fa-basketball-ball" style="color:darkorange;"></i>&nbsp;<a class="blue" href="playerprofile.asp?pid=<%=apid%>"><%=left(apName,14)%></a>&nbsp;<%=apBarps%></td>
							</tr>
						</table>
					</td>	
				<%else%>
					<td class="big" style="text-align:left;vertical-align:middle;width:40%;background-color:black;color:yellow;"><span><i class="fab fa-black-tie"></i>&nbsp;<%=objRBox.Fields("htotBarps").value%>&nbsp;<%=objRBox.Fields("hometeamshort").value%></span>&nbsp;(<%=objrsTeamRecH.Fields("won").Value %>-<%=objrsTeamRecH.Fields("loss").Value %>)</td>
					<td class="big" style="text-align:left;vertical-align:middle;background-color:black;color:yellow;width:40%;"><%=objRBox.Fields("awayteamshort").value%>&nbsp;<%=objRBox.Fields("atotBarps").value%>&nbsp;<i class="fab fa-black-tie"></i></td>
				<%end if%>
				<td  class="big" style="vertical-align: middle;width:10%;"><button type="button" name="box" value="<%=objRBox.Fields("game_Number").value%>" class="btn btn-default btn-block" data-toggle="modal" data-target="#box<%=objRBox.Fields("game_Number").value%>">BOX</button></td>
			</tr>
			<%
			objRBox.MoveNext
			objrsTeamRecA.Close
			objrsTeamRecH.Close
			Wend
			%>
		</table>
<!--MODAL BOX 1-->
<div class="container-fluid">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div id="box1" class="modal fade" role="dialog">
				<div class="modal-dialog" role="document">							
					<div class="modal-content">
						<div class="modal-header modal-header-modal" style="background: yellowgreen;">
							<button type="button" class="close" data-dismiss="modal">&times;</button>
							<h6 class="modal-title">Box Score Matchup #1 </h6>
						</div>
						<% 
						Set objRBox = Server.CreateObject("ADODB.RecordSet")
						objRBox.Open "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"# and game_Number = 1",objConn,3,3,1	
						%>
						<!--#include virtual="Common/indBox.inc"-->
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<!--MODAL BOX 2-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div id="box2" class="modal fade" role="dialog">
				<div class="modal-dialog" role="document">							
					<div class="modal-content">
						<div class="modal-header modal-header-modal" style="background: yellowgreen;">
							<button type="button" class="close" data-dismiss="modal">&times;</button>
							<h6 class="modal-title">Box Score Matchup #2</h6>
						</div>
						<% 
						Set objRBox = Server.CreateObject("ADODB.RecordSet")
						objRBox.Open "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"# and game_Number = 2",objConn,3,3,1	
						%>
						<!--#include virtual="Common/indBox.inc"-->
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<!--MODAL BOX 3-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div id="box3" class="modal fade" role="dialog">
				<div class="modal-dialog" role="document">							
					<div class="modal-content">
						<div class="modal-header modal-header-modal" style="background: yellowgreen;">
							<button type="button" class="close" data-dismiss="modal">&times;</button>
							<h6 class="modal-title">Box Score Matchup #3</h6>
						</div>
						<% 
						Set objRBox = Server.CreateObject("ADODB.RecordSet")
						objRBox.Open "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"# and game_Number = 3",objConn,3,3,1	
						%>
						<!--#include virtual="Common/indBox.inc"-->
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<!--MODAL BOX 4-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div id="box4" class="modal fade" role="dialog">
				<div class="modal-dialog" role="document">							
					<div class="modal-content">
						<div class="modal-header modal-header-modal" style="background: yellowgreen;">
							<button type="button" class="close" data-dismiss="modal">&times;</button>
							<h6 class="modal-title">Box Score Matchup #4</h6>
						</div>
						<% 
						Set objRBox = Server.CreateObject("ADODB.RecordSet")
						objRBox.Open "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"# and game_Number = 4",objConn,3,3,1	
						%>
						<!--#include virtual="Common/indBox.inc"-->
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<!--MODAL BOX 5-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div id="box5" class="modal fade" role="dialog">
				<div class="modal-dialog" role="document">							
					<div class="modal-content">
						<div class="modal-header modal-header-modal" style="background: yellowgreen;">
							<button type="button" class="close" data-dismiss="modal">&times;</button>
							<h6 class="modal-title">Box Score Matchup #5</h6>
						</div>
						<% 
						Set objRBox = Server.CreateObject("ADODB.RecordSet")
						objRBox.Open "SELECT * FROM NEWBOX WHERE gameDate = #"&gameDate&"# and game_Number = 5",objConn,3,3,1	
						%>
						<!--#include virtual="Common/indBox.inc"-->
					</div>
				</div>
			</div>
		</div>
	</div>
</div>
<!--#include virtual="Common/functions.inc"-->