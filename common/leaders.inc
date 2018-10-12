				<%Set objRSLeaders   = Server.CreateObject("ADODB.RecordSet")
					objRSLeaders.Open  "SELECT MAX(gameDate) as nbaGameDate FROM tblLast5 ", objConn,3,3,1
				
					if IsNull(objRSLeaders.Fields("nbaGameDate")) then
						wGameDay = date()
					else
						wGameDay = objRSLeaders.Fields("nbaGameDate").Value
					end if

					objRSLeaders.Close
					objRSLeaders.Open  	"SELECT tblLast5.*, tblPlayers.pid, tblPlayers.ownerid, tblPlayers.pos, TBLOWNERS.shortname, " & _
															"tblNBATeams.teamShortName " & _
															"FROM ((tblLast5 LEFT JOIN tblPlayers ON (tblLast5.Last = tblPlayers.lastName) AND (tblLast5.First = tblPlayers.firstName)) " & _
															"LEFT JOIN TBLOWNERS ON tblPlayers.ownerID = TBLOWNERS.ownerID) " & _
															"LEFT JOIN tblNBATeams ON tblPlayers.NBATeamID = tblNBATeams.NBATID " & _
															"WHERE tblLast5.gameDate= #"&wGameDay&"# " &_
															"ORDER BY tblLast5.barptot DESC, tblLast5.last", objConn,3,3,1
				%>

				<!--
				<table class="table table-custom-black table-responsive table-bordered table-condensed">
					<tr style="background-color:yellowgreen;color:black;font-weight:bold;font-size:12px;">
						<th style="font-size:12px;" colspan="4">BARP LEADERS | <%= wGameDay %></th>
					</tr>
					<tr style="background-color:white;color:black;font-weight:bold;font-size:13px">
						<th class="big" style="width:25%;">Status</th>
						<td class="big" style="width:25%;">Rostered</td>
						<td class="big" style="width:25%;background-color:#fcf8e3;">Available</td>
						<td class="big" style="width:25%;background-color:#dff0d8">My Player</td>
					</tr>
				</table>
				-->
				<table class="table table-striped table-bordered table-custom-black table-condensed" width="100%" class="display" id="example">
					<thead>
						<tr style="background-color:#ddd;font-weight:bold;color:grey;">
							<th class="big" style="vertical-align:middle;text-align:center;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i><span></th>
							<th class="big" style="vertical-align:middle;text-align:left;"><span style="color:black;">PLAYER<span></th>
							<th class="big" style="text-decoration: underline;">B</th>
							<th class="big" style="text-decoration: underline;">A</th>
							<th class="big" style="text-decoration: underline;">R</th>
							<th class="big" style="text-decoration: underline;">P</th>
							<th class="big" style="text-decoration: underline;">S</th>
							<th class="big" style="text-decoration: underline;">3</th>
							<th class="big" style="text-decoration: underline;">T</th>
						</tr>
					</thead>
					<tbody>
					<%
						While Not objRSLeaders.eof
							<!--INITIALIZE 3DBL COUNTERS -->

							tbonus = 0
							'Response.Write "Start Loop is: " & tbonus & ".<br>"

							<!--BONUS INDICATOR CHECK FOR HOME PLAYERS -->

							if Cint(objRSLeaders.Fields("blk").Value) >= 10 then 
								tbonus = tbonus + 1
							end if

							if Cint(objRSLeaders.Fields("ast").Value) >= 10 then 
								tbonus = tbonus + 1
							end if

							if Cint(objRSLeaders.Fields("trb").Value) >= 10 then 
								tbonus = tbonus + 1
							end if

							if Cint(objRSLeaders.Fields("pts").Value) >= 10 then 
								tbonus = tbonus + 1
							end if

							if Cint(objRSLeaders.Fields("stl").Value) >= 10 then 
								tbonus = tbonus + 1
							end if

							if Cint(objRSLeaders.Fields("x3p").Value) >= 10 then 
								tbonus = tbonus + 1
							end if
					%>
					<%if objRSLeaders.Fields("ownerID").Value > 0 then%>
						 <%if objRSLeaders.Fields("ownerID").Value = ownerid then%>
							<tr class="success" style="vertical-align:middle;text-align:center;">		
						 <%else%>
							<tr style="background-color:white;vertical-align:middle;text-align:center;">	
						 <%end if%>	
							<td style="vertical-align:middle;text-align:center;background-color:white;color:black;"><%=round(objRSLeaders.Fields("barptot").Value,2) %></td>									
							<td class="big" style="vertical-align:middle;text-align:left;">
							<a class="blue" href="playerprofile.asp?pid=<%=objRSLeaders.Fields("PID").Value %>"><%=left(objRSLeaders.Fields("first").Value,1)%>.&nbsp;<%=left(objRSLeaders.Fields("last").Value,12)%></a>
							<span class="greenTrade" style="text-transform: uppercase;"><small><%=left(objRSLeaders.Fields("shortname").Value,10)%></span>&nbsp;<span class="orange"><%=objRSLeaders.Fields("pos").Value%></small></span>
							</td>							   							   
					<%else%>
						<tr class="warning">							
						<td style="vertical-align:middle;text-align:center;background-color:white;color:black;"><%=round(objRSLeaders.Fields("barptot").Value,2) %></td>									
						<td class="big" style="vertical-align:middle;text-align:left;">
							<a class="blue" href="playerprofile.asp?pid=<%=objRSLeaders.Fields("PID").Value %>"><%=left(objRSLeaders.Fields("first").Value,1)%>.&nbsp;<%=left(objRSLeaders.Fields("last").Value,12)%></a>
							<small><span class="greenTrade" style="text-transform: uppercase;"><%=left(objRSLeaders.Fields("teamshortname").Value,10)%></span>&nbsp;<span class="orange"><%=objRSLeaders.Fields("pos").Value%></span></small>
						</td>
					<%end if%>
						<% if tbonus >= 3 and CInt(objRSLeaders.Fields("blk").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width:thin;vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("blk").Value)%></td>
						<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("blk").Value)%></td>
						<%end if%>						
						
						<% if tbonus >= 3 and CInt(objRSLeaders.Fields("ast").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("ast").Value)%></td>
						<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("ast").Value)%></td>
						<%end if%>	
	
						<% if tbonus >= 3 and CInt(objRSLeaders.Fields("trb").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("trb").Value)%></td>
						<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("trb").Value)%></td>
						<%end if%>
						
						<% if tbonus >= 3 and CInt(objRSLeaders.Fields("pts").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("pts").Value)%></td>
						<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("pts").Value)%></td>
						<%end if%>

						<% if tbonus >= 3 and CInt(objRSLeaders.Fields("stl").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("stl").Value)%></td>
						<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("stl").Value)%></td>
						<%end if%>						

						<% if tbonus >= 3 and CInt(objRSLeaders.Fields("x3p").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("x3p").Value)%></td>
						<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><%=(objRSLeaders.Fields("x3p").Value)%></td>
						<%end if%>	
							<td class="big" style="vertical-align:middle;text-align:center"><span class="gameTip"><%= objRSLeaders.Fields ("tov").Value %></span></td>
						</tr>
					<%									
						objRSLeaders.MoveNext
						tbonus = 0
						Wend
					%>
					</tbody>
				</table>