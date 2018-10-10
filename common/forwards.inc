<tr>
						<% if objRSForwards.Fields("playerStatus").Value = "W" then %>
						<td align="center">
							<button type="submit" value="Waiver Claim;<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
						</td>
						<% elseif objRSForwards.Fields("playerStatus").Value = "F" and objRSToday.RecordCount > 0 then %>
							<% if objRSNBASked.RecordCount > 0 then %>
							<td style="white-space:nowrap" align="center">
								<button type="submit" value="Sign Free Agent;<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
								<button type="submit" value="Rent Player(s);<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong></button>
							</td>
							<% else %>
							<td align="center">
								<button type="submit" value="Sign Free Agent;<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
							</td>
							<% end if %>
						<% elseif objRSForwards.Fields("playerStatus").Value = "S" then %>
							<% if objRSNBASked.RecordCount > 0 AND wTipTime > (time() - 1/24) then %>
							<td style="white-space:nowrap" align="center">
								<button type="submit" value="Waiver Claim;<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
								<button type="submit" value="Rent Player(s);<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong></button>
							</td>
							<% else %>
							<td align="center">
								<button type="submit" value="Waiver Claim;<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
							</td>
							<% end if %>
						<% else %>
						<td align="center">
							<button type="submit" value="Sign Free Agent;<%=objRSForwards.Fields("PID").Value & ";" & objRSForwards.Fields("firstName").Value & " " & objRSForwards.Fields("lastName").Value & ";" &  objRSForwards.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
						</td>
						<% end if %>

						<% if objRSNBASked.RecordCount > 0 then %>
							 <td><a href="#689<%=objRSForwards.Fields("PID").Value %>" data-toggle="collapse" target="_self"><%=left(objRSForwards.Fields("firstName").Value,1)%>.&nbsp;<%=objRSForwards.Fields("lastName").Value%></a>&nbsp;<orange><%=objRSForwards.Fields("pos").Value%></orange>&nbsp;|&nbsp;<span class="badge" data-toggle="tooltip" title="Season BARP AVG!"><%= round(objRSForwards.Fields("barps").Value,0) %></span>
							<% if CInt(objRSForwards.Fields("l5barps").Value) > CInt(objRSForwards.Fields("barps").Value) then %>
							<i class="fa fa-long-arrow-up" aria-hidden="true"></i>&nbsp;<span class="badgeUp" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSForwards.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSForwards.Fields("teamName").Value %></greenSmall>
							<% elseif CInt(objRSForwards.Fields("barps").Value) > CInt(objRSForwards.Fields("l5barps").Value) then%>
							<i class="fa fa-long-arrow-down" aria-hidden="true"></i>&nbsp;<span class="badgeDown" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSForwards.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSForwards.Fields("teamName").Value %></greenSmall>
							<%else%>
							<i class="fa fa-arrows-h" aria-hidden="true"></i>&nbsp;<span class="badgeEven" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSForwards.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSForwards.Fields("teamName").Value %></greenSmall>
							<%end if %>
							<br><gameTip><%= objRSNBASked.Fields("opponent").Value %>&nbsp;<span class="glyphicon glyphicon-time"></span>&nbsp;<%=wtime%></gameTip>
						<% else %>
						   <td><a href="#689<%=objRSForwards.Fields("PID").Value %>" data-toggle="collapse" target="_self"><%=left(objRSForwards.Fields("firstName").Value,1)%>.&nbsp;<%=objRSForwards.Fields("lastName").Value%></a>&nbsp;<orange><%=objRSForwards.Fields("pos").Value%></orange>&nbsp;|&nbsp;<span class="badge" data-toggle="tooltip" title="Season BARP AVG!"><%= round(objRSForwards.Fields("barps").Value,0) %></span>
							<% if CInt(objRSForwards.Fields("l5barps").Value) > CInt(objRSForwards.Fields("barps").Value) then %>
							<i class="fa fa-long-arrow-up" aria-hidden="true"></i>&nbsp;<span class="badgeUp" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSForwards.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSForwards.Fields("teamName").Value %></greenSmall>
							<% elseif CInt(objRSForwards.Fields("barps").Value) > CInt(objRSForwards.Fields("l5barps").Value) then%>
							<i class="fa fa-long-arrow-down" aria-hidden="true"></i>&nbsp;<span class="badgeDown" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSForwards.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSForwards.Fields("teamName").Value %></greenSmall>
							<%else%>
							<i class="fa fa-arrows-h" aria-hidden="true"></i>&nbsp;<span class="badgeEven" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSForwards.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSForwards.Fields("teamName").Value %></greenSmall>
							<%end if %>
							 <br><%= objRSNBASked.Fields("opponent").Value %>
						<%end if%>		
							 <div id="689<%=objRSForwards.Fields("PID").Value %>" class="collapse">
								<table class="table table-striped table-bordered table-condensed">
									<tr><td style="text-align:center;width:50%"><small>Blocks</small></td><td style="text-align:center;width:50%"><small><%=objRSForwards.Fields("blk").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Assists</small></td><td style="text-align:center;width:50%"><small><%=objRSForwards.Fields("ast").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Rebounds</small></td><td style="text-align:center;width:50%"><small><%=objRSForwards.Fields("reb").Value %></td></small></tr>
									<tr><td style="text-align:center;width:50%"><small>Points</small></td><td style="text-align:center;width:50%"><small><%=objRSForwards.Fields("ppg").Value %></small></td></tr>	
									<tr><td style="text-align:center;width:50%"><small>Steals</small></td><td style="text-align:center;width:50%"><small><%=objRSForwards.Fields("stl").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>3-Pointers</small></td><td style="text-align:center;width:50%"><small><%=objRSForwards.Fields("three").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Turnovers</small></td><td style="text-align:center;width:50%"><small><%=objRSForwards.Fields("to").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Barps</small></td><td style="text-align:center;width:50%"><small><%= objRSForwards.Fields("barps").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Clear Date</small></td><td style="text-align:center;width:50%"><small><%= objRSForwards.Fields("clearwaiverdate").Value %></small></td></tr>
								</table>
							</div>							 
							</td>							 
					</tr>