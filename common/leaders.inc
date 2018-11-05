				<%Set objRSLeaders   = Server.CreateObject("ADODB.RecordSet")
					objRSLeaders.Open  "SELECT MAX(gameDate) as nbaGameDate FROM tblLast5 ", objConn,3,3,1
				
					if IsNull(objRSLeaders.Fields("nbaGameDate")) then
						wGameDay = date()
					else
						wGameDay = objRSLeaders.Fields("nbaGameDate").Value
					end if

					objRSLeaders.Close
					
					'Response.Write "Stats for: "&wGameDay&"<br>"
					
					objRSLeaders.Open  	"SELECT tblLast5.*, tblPlayers.pid, tblPlayers.ownerid, tblPlayers.pos, TBLOWNERS.shortname, " & _
															"tblNBATeams.teamShortName " & _
															"FROM ((tblLast5 LEFT JOIN tblPlayers ON (tblLast5.Last = tblPlayers.lastName) AND (tblLast5.First = tblPlayers.firstName)) " & _
															"LEFT JOIN TBLOWNERS ON tblPlayers.ownerID = TBLOWNERS.ownerID) " & _
															"LEFT JOIN tblNBATeams ON tblPlayers.NBATeamID = tblNBATeams.NBATID " & _
															"WHERE tblLast5.gameDate= #"&wGameDay&"# " &_
															"ORDER BY tblLast5.barptot DESC, tblLast5.last", objConn,3,3,1
				%>
				</br>
				<div class="row">		
					<div class="col-md-12 col-sm-12 col-xs-12">			
						<span style="font-size:12px;color:black;" class="pull-right"><strong><%=FormatDateTime(wGameDay,1)%></strong></span></br>
					</div>
				</div>
				</br>
				<table class="table table-striped table-bordered table-custom-black table-condensed" width="100%" class="display" id="example1">
					<thead>
						<tr style="background-color:#ddd;font-weight:bold;color:grey;">
							<th class="big" style="vertical-align:middle;text-align:center;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i><span></th>
							<th class="big" style="vertical-align:middle;text-align:left;"><span style="color:black;">PLAYER<span></th>
							<th class="big hidden-xs" style="text-decoration: underline;">M</th>
						  <th class="big hidden-xs" style="text-decoration: underline;">U</th>
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
							
							if isnull(objRSLeaders.Fields("ownerID").Value) then
							   wPID      = 9994
				               wTeam     = "NA"
				               wPOS      = "NA"							
							else
							   wPID      = objRSLeaders.Fields("PID").Value
				               wTeam     = objRSLeaders.Fields("teamShortName").Value
				               wPOS      = objRSLeaders.Fields("pos").Value
							end if
					%>
					<%if objRSLeaders.Fields("ownerID").Value > 0 then%>
						 <%if objRSLeaders.Fields("ownerID").Value = ownerid then%>
							<tr class="success" style="vertical-align:middle;text-align:center;">		
						 <%else%>
							<tr style="background-color:white;vertical-align:middle;text-align:center;">	
						 <%end if%>	
							<td style="vertical-align:middle;text-align:center;color:black;font-weight:bold;"><%=round(objRSLeaders.Fields("barptot").Value,2) %></td>									
							<td class="big" style="vertical-align:middle;text-align:left;">
							<a class="blue" href="playerprofile.asp?pid=<%=wPID %>"><%=left(objRSLeaders.Fields("first").Value,1)%>.&nbsp;<%=left(objRSLeaders.Fields("last").Value,11)%></a>
							<span class="greenTrade" style="text-transform: uppercase;"><%=left(wTeam,10)%></span>&nbsp;<span class="orange"><%=wPOS%></span>
							</td>							   							   
					<%else%>
						<tr class="warning">							
						<td style="vertical-align:middle;text-align:center;font-weight:bold;"><%=round(objRSLeaders.Fields("barptot").Value,2) %></td>									
						<td class="big" style="vertical-align:middle;text-align:left;">
							<a class="blue" href="playerprofile.asp?pid=<%=wPID %>"><%=left(objRSLeaders.Fields("first").Value,1)%>.&nbsp;<%=left(objRSLeaders.Fields("last").Value,12)%></a>&nbsp;<small style="color:red;font-weight:bold;">[FA]</small>
							<span class="greenTrade" style="text-transform: uppercase;"><%=left(wTeam,10)%></span>&nbsp;<span class="orange"><%=wPOS%></span>
						</td>
					<%end if%>
						<td class="big hidden-xs" style="vertical-align: middle;"><%=objRSLeaders.Fields("MP").Value%></td>
						<td class="big hidden-xs" style="vertical-align: middle;"><%=objRSLeaders.Fields("usage").Value%></td>		
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
				</br>