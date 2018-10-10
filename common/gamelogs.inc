				<% 
					Set objRSLogs  = Server.CreateObject("ADODB.RecordSet")
					objRSLogs.Open   "SELECT * FROM qryResultHistory where HomeTeamID = "&ownerid&"" , objConn,3,3,1
				%>

				<table class="table table-custom-black table-responsive table-bordered table-condensed">
					<thead>
						<tr style="background-color:yellowgreen;color:black;font-weight:bold;">
							<th class="big" width="20%">DATE</th>
							<th class="big">OPPONENT</th>
							<th class="big" width="15%" colspan="3" style="text-align:center">RESULTS</td>
						</tr>
					</thead>
					<tbody>
					<%
						While Not objRSLogs.eof
						wTeamName = replace(UCase(objRSLogs.Fields("Villian_Team").Value), "THE ", "")
					%>
					<%if objRSLogs.Fields ("Result").Value = "W" then %>
						<tr class="success">	
					<%else %>
						<tr bgcolor="#FFFFFF">	
					<%end if %>										
							<td class="big" style="vertical-align:middle;text-align:left"><a class="blue" href="box_master.asp?gameDate=<%=objRSLogs.Fields("gameDate").Value %>"><%= objRSLogs.Fields("GameDate").Value %></a></td>
					<% if len(wTeamName) >= 20 then %>
							<td class="big" style="vertical-align:middle;text-align:left"><span class="gameTip"><%=objRSLogs.Fields ("villian_team_short").Value%></span></td>
					<% else %>
							<td class="big" style="vertical-align:middle;text-align:left"><span class="gameTip"><%=wTeamName%></span></td>					
					<% end if %>
					
					<!--
					<%if objRSLogs.Fields ("Result").Value = "W" and objRSLogs.Fields ("Penalty").Value = true then %>
							<td class="big" width="10%" style="vertical-align:middle;text-align:center"><span class="gameTip">P</span> <span class="redTrade"><i class="fa fa-asterisk" aria-hidden="true"></i></span></td>
					<%else%>
							<td class="big" width="10%" style="vertical-align:middle;text-align:center"><span class="gameTip"><%= objRSLogs.Fields ("Result").Value %></span></td>
					<%end if%>
					<%if objRSLogs.Fields ("Result").Value = "W" then %>
							<td class="big" style="vertical-align:middle;text-align:center;font-weight:bold;" width="10%"><span class="greenTrade"><%= objRSLogs.Fields ("Hero_Score").Value %></span></td>
					<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;font-weight:bold;" width="10%"><span class="redTrade"><%= objRSLogs.Fields ("Hero_Score").Value %></span></td>
					<%end if%>
							<td class="big" style="vertical-align:middle;text-align:center" width="10%"><span class="gameTip"><%= objRSLogs.Fields ("Villian_Score").Value %></span></td>
						</tr>
						-->
					
					<%if objRSLogs.Fields ("Result").Value = "W" and objRSLogs.Fields ("Penalty").Value = true then %>
							<td class="big" width="10%" style="vertical-align:middle;text-align:center"><span class="gameTip">P</span> <span class="redTrade"><i class="fa fa-asterisk" aria-hidden="true"></i>&nbsp;<%= objRSLogs.Fields ("Result").Value %></span></span></td>
							<td class="big" style="vertical-align:middle;text-align:center;font-weight:bold;" width="10%"><span class="greenTrade"><%= objRSLogs.Fields ("Hero_Score").Value %></span></td>
							<td class="big" style="vertical-align:middle;text-align:center" width="10%"><span class="gameTip"><%= objRSLogs.Fields ("Villian_Score").Value %></span></td>
					<%elseif objRSLogs.Fields ("Result").Value = "W" then %>
							<td class="big" width="10%" style="vertical-align:middle;text-align:center;font-weight:bold;"><span class="greenTrade"><%= objRSLogs.Fields ("Result").Value %></span></td>
							<td class="big" width="10%" style="vertical-align:middle;text-align:center;font-weight:bold;"><span class="greenTrade"><%= objRSLogs.Fields ("Hero_Score").Value %></span></td>
							<td class="big" width="10%" style="vertical-align:middle;text-align:center;font-weight:bold;color:black;"><%= objRSLogs.Fields ("Villian_Score").Value %></td>
					<%else%>
						<td class="big" width="10%" style="vertical-align:middle;text-align:center;font-weight:bold;"><span class="redTrade"><%= objRSLogs.Fields ("Result").Value %></span></td>
						<td class="big" width="10%" style="vertical-align:middle;text-align:center;font-weight:bold;"><span class="redTrade"><%= objRSLogs.Fields ("Hero_Score").Value %></span></td>
						<td class="big" width="10%" style="vertical-align:middle;text-align:center;font-weight:bold;color:black;"><%= objRSLogs.Fields ("Villian_Score").Value %></td>
					<%end if%>

						<%
						objRSLogs.MoveNext
						Wend
					%>
						<tr style="background-color:#424242;color:white;font-weight:bold;">
							<td colspan="5"><strong><span class="redTrade"><i class="fa fa-asterisk redTrade" aria-hidden="true"></i></span> denotes loss awarded due to penalty lineup used.</strong></td>
						</tr>
					</tbody>
				</table>