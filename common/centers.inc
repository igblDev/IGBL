					<tr>
						<% if objRSCenters.Fields("playerStatus").Value = "W" then %>
						<td align="center">
							<button type="submit" value="Waiver Claim;<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
						</td>
						<% elseif objRSCenters.Fields("playerStatus").Value = "F" and objRSToday.RecordCount > 0 then %>
							<% if objRSNBASked.RecordCount > 0 then %>
							<td style="white-space:nowrap" align="center">
								<button type="submit" value="Sign Free Agent;<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
								<button type="submit" value="Rent Player(s);<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong></button>
							</td>
							<% else %>
							<td align="center">
								<button type="submit" value="Sign Free Agent;<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
							</td>
							<% end if %>
						<% elseif objRSCenters.Fields("playerStatus").Value = "S" then %>
							<% if objRSNBASked.RecordCount > 0 AND wTipTime > (time() - 1/24) then %>
							<td style="white-space:nowrap" align="center">
								<button type="submit" value="Waiver Claim;<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
								<button type="submit" value="Rent Player(s);<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong></button>
							</td>
							<% else %>
							<td align="center">
								<button type="submit" value="Waiver Claim;<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
							</td>
							<% end if %>
						<% else %>
						<td align="center">
							<button type="submit" value="Sign Free Agent;<%=objRSCenters.Fields("PID").Value & ";" & objRSCenters.Fields("firstName").Value & " " & objRSCenters.Fields("lastName").Value & ";" &  objRSCenters.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
						</td>
						<% end if %>

						<% if objRSNBASked.RecordCount > 0 then %>
							 <td><a href="#350<%=objRSCenters.Fields("PID").Value %>" data-toggle="collapse" target="_self"><%=left(objRSCenters.Fields("firstName").Value,1)%>.&nbsp;<%=objRSCenters.Fields("lastName").Value%></a>&nbsp;<orange><%=objRSCenters.Fields("pos").Value%></orange>&nbsp;|&nbsp;<span class="badge" data-toggle="tooltip" title="Season BARP AVG!"><%= round(objRSCenters.Fields("barps").Value,0) %></span>
							<% if CInt(objRSCenters.Fields("l5barps").Value) > CInt(objRSCenters.Fields("barps").Value) then %>
							<i class="fa fa-long-arrow-up" aria-hidden="true"></i>&nbsp;<span class="badgeUp" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSCenters.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSCenters.Fields("teamName").Value %></greenSmall>
							<% elseif CInt(objRSCenters.Fields("barps").Value) > CInt(objRSCenters.Fields("l5barps").Value) then%>
							<i class="fa fa-long-arrow-down" aria-hidden="true"></i>&nbsp;<span class="badgeDown" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSCenters.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSCenters.Fields("teamName").Value %></greenSmall>
							<%else%>
							<i class="fa fa-arrows-h" aria-hidden="true"></i>&nbsp;<span class="badgeEven" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSCenters.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSCenters.Fields("teamName").Value %></greenSmall>
							<%end if %>
							<br><gameTip><%= objRSNBASked.Fields("opponent").Value %>&nbsp;<span class="glyphicon glyphicon-time"></span>&nbsp;<%=wtime%></gameTip>
						<% else %>
						   <td><a href="#350<%=objRSCenters.Fields("PID").Value %>" data-toggle="collapse" target="_self"><%=left(objRSCenters.Fields("firstName").Value,1)%>.&nbsp;<%=objRSCenters.Fields("lastName").Value%></a>&nbsp;<orange><%=objRSCenters.Fields("pos").Value%></orange>&nbsp;|&nbsp;<span class="badge" data-toggle="tooltip" title="Season BARP AVG!"><%= round(objRSCenters.Fields("barps").Value,0) %></span>
							<% if CInt(objRSCenters.Fields("l5barps").Value) > CInt(objRSCenters.Fields("barps").Value) then %>
							<i class="fa fa-long-arrow-up" aria-hidden="true"></i>&nbsp;<span class="badgeUp" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSCenters.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSCenters.Fields("teamName").Value %></greenSmall>
							<% elseif CInt(objRSCenters.Fields("barps").Value) > CInt(objRSCenters.Fields("l5barps").Value) then%>
							<i class="fa fa-long-arrow-down" aria-hidden="true"></i>&nbsp;<span class="badgeDown" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSCenters.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSCenters.Fields("teamName").Value %></greenSmall>
							<%else%>
							<i class="fa fa-arrows-h" aria-hidden="true"></i>&nbsp;<span class="badgeEven" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSCenters.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSCenters.Fields("teamName").Value %></greenSmall>
							<%end if %>
							 <br><%= objRSNBASked.Fields("opponent").Value %>
						<%end if%>		
							 <div id="350<%=objRSCenters.Fields("PID").Value %>" class="collapse">
								<table class="table table-striped table-bordered table-condensed">
									<tr><td style="text-align:center;width:50%"><small>Blocks</small></td><td style="text-align:center;width:50%"><small><%=objRSCenters.Fields("blk").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Assists</small></td><td style="text-align:center;width:50%"><small><%=objRSCenters.Fields("ast").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Rebounds</small></td><td style="text-align:center;width:50%"><small><%=objRSCenters.Fields("reb").Value %></td></small></tr>
									<tr><td style="text-align:center;width:50%"><small>Points</small></td><td style="text-align:center;width:50%"><small><%=objRSCenters.Fields("ppg").Value %></small></td></tr>	
									<tr><td style="text-align:center;width:50%"><small>Steals</small></td><td style="text-align:center;width:50%"><small><%=objRSCenters.Fields("stl").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>3-Pointers</small></td><td style="text-align:center;width:50%"><small><%=objRSCenters.Fields("three").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Turnovers</small></td><td style="text-align:center;width:50%"><small><%=objRSCenters.Fields("to").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Barps</small></td><td style="text-align:center;width:50%"><small><%= objRSCenters.Fields("barps").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Clear Date</small></td><td style="text-align:center;width:50%"><small><%= objRSCenters.Fields("clearwaiverdate").Value %></small></td></tr>
								</table>
							</div>							 
							</td>							 
					</tr>	