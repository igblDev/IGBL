<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<%
	On Error Resume Next
	Session("FP_OldCodePage") = Session.CodePage
	Session("FP_OldLCID") = Session.LCID
	Session.CodePage = 1252
	Err.Clear
	strErrorUrl = ""

	Dim objConn, sTeam, sAction, sURL,ownerid


	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

	%>
	<!--#include virtual="Common/session.inc"-->
	<%	

	ownerid = session("ownerid")
								
	%>
<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<!--#include virtual="Common/headerMain.inc"-->
<style>
darkorange {
	color: darkorange;
}

th {
    vertical-align: middle;
  	text-align: center;
}

td {
    vertical-align: middle;
    text-align: left;
}

black {
	color:black;
	text-transform: uppercase;
}
.nav-tabs {
    border-bottom: 2px solid black;
}

</style>
</head>
<body>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Schedules</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<ul class="nav nav-tabs">
				<li class="active"><a data-toggle="tab" href="#mysked"><i class="fal fa-calendar-alt"></i>&nbsp;My Schedule </a></li>
				<li><a data-toggle="tab" href="#leadersked"><i class="fas fa-calendar-alt"></i>&nbsp;League Schedule</a></li>
				</ul>
		</div>
	</div>
	</br>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="tab-content">
				<div id="mysked" class="tab-pane fade in active">
					<%	
					ownerid = session("ownerid")
					Dim objRSMyGames, gameCnt												
					Set objRSMyGames	= Server.CreateObject("ADODB.RecordSet")
					objRSMyGames.Open  	"SELECT * FROM qryAllGames where AwayTeamInd = " & ownerid & " " &_
										"OR HomeTeamInd  = " & ownerid & " ", objConn		

											
					%>
					<table class="table table-custom-black table-responsive table-bordered table-condensed">
						<tr style="background-color:yellowgreen;color:black;font-weight:bold;">
							<th class="big">Date</th>
							<th class="big">Opponent</th>
							<th class="big">Tip <small>[cst]</small></th>
							<th class="big">C</th>								
						</tr>
						<% While Not objRSMyGames.EOF %>
						<%     
							if objRSMyGames("HomeTeamInd") = ownerid then
								OpponentName = objRSMyGames("AwayTeamName")
								gameloc = "vs."
							else
								OpponentName = objRSMyGames("HomeTeamName")
								gameloc = "vs."
							end if
						  if len(objRSMyGames.Fields("gamedeadline").Value) = 10 then
								wtime = Left(objRSMyGames.Fields("gamedeadline").Value,4) & Right(objRSMyGames.Fields("gamedeadline").Value,3)
						  else
								wtime = Left(objRSMyGames.Fields("gamedeadline").Value,5) & Right(objRSMyGames.Fields("gamedeadline").Value,3)
						  end if							
						%>

						 <%if objRSMyGames("cycle").Value = 1 then%>
								<tr style="background-color:white;">	
						 <%elseif objRSMyGames("cycle").Value = 2 then%>
								<tr style="background-color:black;color:white;">	
						 <%elseif objRSMyGames("cycle").Value = 3 then%>
								<tr style="background-color:yellowgreen;color:white;">	
						 <%elseif objRSMyGames("cycle").Value = 4 then%>
								<tr style="background-color:white;">	
						 <%elseif objRSMyGames("cycle").Value = 5 then%>						 
								<tr style="background-color:black;color:white;">	
						 <%elseif objRSMyGames("cycle").Value = 6 then%>
								<tr style="background-color:yellowgreen;color:white;">	
						 <%elseif objRSMyGames("cycle").Value = 7 then%>
								<tr style="background-color:white;">	
						 <%end if%>									
								<td class="big text-center"><%=objRSMyGames("gameday")%></td>
								<td class="big"><%=gameloc%>&nbsp;<strong><%=OpponentName%></strong></td>
								<td class="big text-center"><%=wtime%></td>
								<td class="big text-center"><%=objRSMyGames("cycle")%></td>
							</tr>

						<%
							objRSMyGames.MoveNext
							Wend
						%>
					</table>				
				</div>
				<div id="leadersked" class="tab-pane">
				<%
					Dim objRSAllDates, objRSDay, objRSPlayoffs
	
					Set objRSAllDates = Server.CreateObject("ADODB.RecordSet")
					Set objRSDay      = Server.CreateObject("ADODB.RecordSet")
					Set objRSPlayoffs = Server.CreateObject("ADODB.RecordSet")
	
					objRSPlayoffs.Open "SELECT param_date FROM tblParameterCtl where param_name = 'PLAYOFF_START_DATE' ", objConn	
					wQtrDate = objRSPlayoffs.Fields("param_date").Value
					objRSPlayoffs.Close
					
					objRSPlayoffs.Open "SELECT param_date FROM tblParameterCtl where param_name = 'PO_SEMIS' ", objConn	
					wSemiDate = objRSPlayoffs.Fields("param_date").Value
					objRSPlayoffs.Close
					
					objRSPlayoffs.Open "SELECT param_date FROM tblParameterCtl where param_name = 'PO_FINALS' ", objConn	
					wFinalDate = objRSPlayoffs.Fields("param_date").Value
					objRSPlayoffs.Close
					
					objRSAllDates.Open "SELECT distinct gameday, gameDeadline FROM qry_matchups ORDER BY GameDay", objConn				
					While Not objRSAllDates.EOF
						wdate = objRSAllDates.Fields("gameday").Value					   
						if cdate(wdate) >= cdate(wFinalDate) then
						  sTeamDisplay = "Finalist"
						elseif cdate(wdate) >= cdate(wSemiDate) then
						  sTeamDisplay = "SemiFinalist"
						elseif cdate(wdate) >= cdate(wQtrDate) then
						  sTeamDisplay = "QuarterFinalist"
						else
						  sTeamDisplay = "### INVALID TEAMS SHOULD NOT BE NULL FOR THIS DATE ###"
						end if	
					 
					 if len(objRSAllDates.Fields("gamedeadline").Value) = 10 then
								wtime = Left(objRSAllDates.Fields("gamedeadline").Value,4) & Right(objRSAllDates.Fields("gamedeadline").Value,3)
					 else
								wtime = Left(objRSAllDates.Fields("gamedeadline").Value,5) & Right(objRSAllDates.Fields("gamedeadline").Value,3)
					 end if								
				%>				
				<table class="table table-custom-black table-responsive table-bordered table-condensed">
					<tr style="background-color:yellowgreen;color:black;font-weight:bold;">
						<th class="text-center big" colspan="3"><%= (FormatDateTime(objRSAllDates.Fields("gameday").Value,1)) %> <span class="glyphicon glyphicon-time"></span> <%= wtime %> <small>[cst]</small></th></th>
					</tr>
				<%
				    objRSDay.Open "SELECT * FROM qry_matchups where GameDay = #"&wdate&"# order by 5", objConn
				    While Not objRSDay.EOF
				%>
					 <tr style="background-color:white;">		
						<% if ISNULL(objRSDay.Fields("Home Team Owner").Value) then %>
							<td class="big"style="text-align:right;width:45%"><greenName><%=sTeamDisplay%></greenName></td>												
						<% elseif objRSDay.Fields("Home Team Owner").Value = ownerid then %>						
							<td class="success big" style="text-align:right;width:45%"><strong><a class="blue" href="dashboard.asp?ownerid=<%= ownerid %>&Action=Retrieve Lineup&currentDate=<%= objRSDay.Fields("gameday").Value%>"> <%= objRSDay.Fields("Home Team Short").Value %></a></strong>
							<%if objRSDay.Fields("HTRank").Value > 6 then%>
							<span style="color: #9a1400;"><i class="far fa-angry"></i></span>
							<%elseif objRSDay.Fields("HTRank")= 1 or objRSDay.Fields("HTRank")= 2 then%>
							<span style="color:darkorange;font-weight:bold;"><%= objRSDay.Fields("HTRank").Value %></span>
							<%elseif objRSDay.Fields("HTRank")= 3 or objRSDay.Fields("HTRank")= 4 or objRSDay.Fields("HTRank")= 5 or  objRSDay.Fields("HTRank")= 6 then%>
							<span style="color:#468847;font-weight:bold;"><%= objRSDay.Fields("HTRank").Value %></span>
							<%else%>
							<span><%= objRSDay.Fields("HTRank").Value %></span>
							<%end if%>			
							</td>
						<% else %>
							<td class="big" style="text-align:right;width:45%"><greenName><%= objRSDay.Fields("Home Team Short").Value %></greenName>
							<%if objRSDay.Fields("HTRank").Value > 6 then%>
							<span style="color: #9a1400;"><i class="far fa-angry"></i></span>	
							<%elseif objRSDay.Fields("HTRank")= 1 or objRSDay.Fields("HTRank")= 2 then%>
							<span style="color:darkorange;font-weight:bold;"><%= objRSDay.Fields("HTRank").Value %></span>
							<%elseif objRSDay.Fields("HTRank")= 3 or objRSDay.Fields("HTRank")= 4 or objRSDay.Fields("HTRank")= 5 or  objRSDay.Fields("HTRank")= 6 then%>
							<span style="color:#468847;font-weight:bold;"><%= objRSDay.Fields("HTRank").Value %></span>
							<%else%>
							<span><%= objRSDay.Fields("HTRank").Value %></span>
							<%end if%>			
							</td>				
						<%end if%>
							<td class="big" style="text-align:center;" width="10%">vs.</td>	
						<% if ISNULL(objRSDay.Fields("Visiting Team Owner").Value) then %>	
							<td class="big" width="45%"><greenName><%=sTeamDisplay%></greenName></td>												
						<% elseif objRSDay.Fields("Visiting Team Owner").Value = ownerid then %>
							<td class="success big" style="width:45%">
							<%if objRSDay.Fields("VTRank").Value > 6 then%>
							<span style="color: #9a1400;"><i class="far fa-angry"></i></span>
							<%elseif objRSDay.Fields("VTRank")= 1 or objRSDay.Fields("VTRank")= 2 then%>
							<span style="color:darkorange;font-weight:bold;"><%= objRSDay.Fields("VTRank").Value %></span>
							<%elseif objRSDay.Fields("VTRank")= 3 or objRSDay.Fields("VTRank")= 4 or objRSDay.Fields("VTRank")= 5 or  objRSDay.Fields("VTRank")= 6 then%>
							<span style="color:#468847;font-weight:bold;"><%= objRSDay.Fields("VTRank").Value %></span>
							<%else%>
							<span><%= objRSDay.Fields("VTRank").Value %></span>
							<%end if%>	
							<strong><a class="blue" href="dashboard.asp?ownerid=<%= ownerid %>&Action=Retrieve Lineup&currentDate=<%= objRSDay.Fields("gameday").Value%>"> <%= objRSDay.Fields("Visiting Team Short").Value %></a></strong>
						</td>														  
						<% else %>
							<td class="big" width="45%"><greenName>
							<%if objRSDay.Fields("VTRank").Value > 6 then%>
							<span style="color: #9a1400;"><i class="far fa-angry"></i></span>
							<%elseif objRSDay.Fields("VTRank")= 1 or objRSDay.Fields("VTRank")= 2 then%>
							<span style="color:darkorange;font-weight:bold;"><%= objRSDay.Fields("VTRank").Value %></span>
							<%elseif objRSDay.Fields("VTRank")= 3 or objRSDay.Fields("VTRank")= 4 or objRSDay.Fields("VTRank")= 5 or  objRSDay.Fields("VTRank")= 6 then%>
							<span style="color:#468847;font-weight:bold;"><%= objRSDay.Fields("VTRank").Value %></span>
							<%else%>
							<span><%= objRSDay.Fields("VTRank").Value %></span>
							<%end if%>
							<%= objRSDay.Fields("Visiting Team Short").Value %></greenName>
							</td>
						<%end if%>
					</tr>
					<%
						objRSDay.MoveNext
							Wend
						objRSDay.Close					
					%>
				</table>
				<br>
				<%					
					objRSAllDates.MoveNext
						Wend 
					objRSAllDates.Close					
				%>				
				</div>
			</div>	
		</div>		
	</div>
</div>	
<% 
objRSMyGames.Close
objConn.Close
objrs.close
objrsgames.close
Set objrs = Nothing
Set objConn = Nothing
 %>
</body>
</html>
