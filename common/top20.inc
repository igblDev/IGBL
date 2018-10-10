					<tr>
						<% if objRSTop20.Fields("playerStatus").Value = "W" then %>
						<td style="white-space:nowrap" align="center">
							<button type="submit" value="Waiver Claim;<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
						</td>
						<% elseif objRSTop20.Fields("playerStatus").Value = "F" and objRSToday.RecordCount > 0 then %>
							<% if objRSNBASked.RecordCount > 0 then %>
							<td style="white-space:nowrap" align="center">
								<button type="submit" value="Sign Free Agent;<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
								<button type="submit" value="Rent Player(s);<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong></button>
							</td>
							<% else %>
							<td align="center">
								<button type="submit" value="Sign Free Agent;<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
							</td>
							<% end if %>
						<% elseif objRSTop20.Fields("playerStatus").Value = "S" then %>
							<% if objRSNBASked.RecordCount > 0 AND wTipTime > (time() - 1/24) then %>
							<td style="white-space:nowrap" align="center">
								<button type="submit" value="Waiver Claim;<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
								<button type="submit" value="Rent Player(s);<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong></button>
							</td>
							<% else %>
							<td align="center">
								<button type="submit" value="Waiver Claim;<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><red><i class="fa fa-plus-square red" aria-hidden="true"></i></red></button>
							</td>
							<% end if %>
						<% else %>
						<td align="center">
							<button type="submit" value="Sign Free Agent;<%=objRSTop20.Fields("PID").Value & ";" & objRSTop20.Fields("firstName").Value & " " & objRSTop20.Fields("lastName").Value & ";" &  objRSTop20.Fields("StatusDesc").Value%>" name="Action" class="btn btn-default btn-xs"><i class="fa fa-plus green" aria-hidden="true"></i></button>
						</td>
						<% end if %>
						<% if objRSNBASked.RecordCount > 0 then %>
               <td><a href="#777<%=objRSTop20.Fields("PID").Value %>" data-toggle="collapse" target="_self"><%=left(objRSTop20.Fields("firstName").Value,1)%>.&nbsp;<%=objRSTop20.Fields("lastName").Value%></a>&nbsp;<orange><%=objRSTop20.Fields("pos").Value%></orange>&nbsp;|&nbsp;<span class="badge" data-toggle="tooltip" title="Season BARP AVG!"><%= round(objRSTop20.Fields("barps").Value,0) %></span>
							<% if CInt(objRSTop20.Fields("l5barps").Value) > CInt(objRSTop20.Fields("barps").Value) then %>
							<i class="fa fa-long-arrow-up" aria-hidden="true"></i>&nbsp;<span class="badgeUp" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSTop20.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSTop20.Fields("teamName").Value %></greenSmall>
							<% elseif CInt(objRSTop20.Fields("barps").Value) > CInt(objRSTop20.Fields("l5barps").Value) then%>
							<i class="fa fa-long-arrow-down" aria-hidden="true"></i>&nbsp;<span class="badgeDown" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSTop20.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSTop20.Fields("teamName").Value %></greenSmall>
							<%else%>
							<i class="fa fa-arrows-h" aria-hidden="true"></i>&nbsp;<span class="badgeEven" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSTop20.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSTop20.Fields("teamName").Value %></greenSmall>
							<%end if %>
							<br><gameTip><%= objRSNBASked.Fields("opponent").Value %>&nbsp;<span class="glyphicon glyphicon-time"></span>&nbsp;<%=wtime%></gameTip>
						<% else %>
						   <td><a href="#777<%=objRSTop20.Fields("PID").Value %>" data-toggle="collapse" target="_self"><%=left(objRSTop20.Fields("firstName").Value,1)%>.&nbsp;<%=objRSTop20.Fields("lastName").Value%></a>&nbsp;<orange><%=objRSTop20.Fields("pos").Value%></orange>&nbsp;|&nbsp;<span class="badge" data-toggle="tooltip" title="Season BARP AVG!"><%= round(objRSTop20.Fields("barps").Value,0) %></span>
							<% if CInt(objRSTop20.Fields("l5barps").Value) > CInt(objRSTop20.Fields("barps").Value) then %>
							<i class="fa fa-long-arrow-up" aria-hidden="true"></i>&nbsp;<span class="badgeUp" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSTop20.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSTop20.Fields("teamName").Value %></greenSmall>
							<% elseif CInt(objRSTop20.Fields("barps").Value) > CInt(objRSTop20.Fields("l5barps").Value) then%>
							<i class="fa fa-long-arrow-down" aria-hidden="true"></i>&nbsp;<span class="badgeDown" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSTop20.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSTop20.Fields("teamName").Value %></greenSmall>
							<%else%>
							<i class="fa fa-arrows-h" aria-hidden="true"></i>&nbsp;<span class="badgeEven" data-toggle="tooltip" title="Last 5 BARP AVG!"><%= objRSTop20.Fields ("l5barps").Value %></span></br><greenSmall><%= objRSTop20.Fields("teamName").Value %></greenSmall>
							<%end if %>
							 <br><%= objRSNBASked.Fields("opponent").Value %> 
						<%end if%>		
							 <div id="777<%=objRSTop20.Fields("PID").Value %>" class="collapse">
								<table class="table table-striped table-bordered table-condensed">
									<tr><td style="text-align:center;width:50%"><small>Blocks</small></td><td style="text-align:center;width:50%"><small><%=objRSTop20.Fields("blk").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Assists</small></td><td style="text-align:center;width:50%"><small><%=objRSTop20.Fields("ast").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Rebounds</small></td><td style="text-align:center;width:50%"><small><%=objRSTop20.Fields("reb").Value %></td></small></tr>
									<tr><td style="text-align:center;width:50%"><small>Points</small></td><td style="text-align:center;width:50%"><small><%=objRSTop20.Fields("ppg").Value %></small></td></tr>	
									<tr><td style="text-align:center;width:50%"><small>Steals</small></td><td style="text-align:center;width:50%"><small><%=objRSTop20.Fields("stl").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>3-Pointers</small></td><td style="text-align:center;width:50%"><small><%=objRSTop20.Fields("three").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Turnovers</small></td><td style="text-align:center;width:50%"><small><%=objRSTop20.Fields("to").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Barps</small></td><td style="text-align:center;width:50%"><small><%= objRSTop20.Fields("barps").Value %></small></td></tr>
									<tr><td style="text-align:center;width:50%"><small>Clear Date</small></td><td style="text-align:center;width:50%"><small><%= objRSTop20.Fields("clearwaiverdate").Value %></small></td></tr>
								</table>
							</div>							 
						</td>							 
					</tr>	