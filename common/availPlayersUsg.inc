							<td class="big" style="vertical-align:middle;text-align:center;vertical-align:bottom;width:20%;padding: inherit;"><img class="img-fluid center-block" src="<%= (objsrAvailPlayers.Fields("image").Value)%>"></td>
							<td class="big" style="vertical-align:middle;text-align:left;width:34%;">
							<%if (len(objsrAvailPlayers.Fields("firstName").Value) + len(objsrAvailPlayers.Fields("lastName").Value)) >= 17 then %>
								<a class="blue" href="playerprofile.asp?pid=<%=objsrAvailPlayers.Fields("PID").Value %>">
									<%=left(objsrAvailPlayers.Fields("firstName").Value,1)%>.&nbsp;<%=left(objsrAvailPlayers.Fields("lastName").Value,14)%></a>
							<%else%>
								<a class="blue" href="playerprofile.asp?pid=<%=objsrAvailPlayers.Fields("PID").Value %>">
									<%=objsrAvailPlayers.Fields("firstName").Value%>&nbsp;<%=objsrAvailPlayers.Fields("lastName").Value%></a>
							<%end if%>		
							<br><span class="greenTrade"><%= objsrAvailPlayers.Fields("team")%></span>&nbsp;<span class="orange"><%= objsrAvailPlayers.Fields("pos")%></span>	
							<% if objsrAvailPlayers.Fields("rentalPlayer").Value = true then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
							<%end if%>
							<%if objsrAvailPlayers.Fields("IR").Value = true then %>
								<strong>	<i class="fas fa-briefcase-medical red"></i></strong>
							<%end if%>	
							<% if objsrAvailPlayers.Fields("pendingtrade").Value = true then %>
								<strong>	<i class="far fa-exchange"></i></strong>
							<%end if%>
							<% if objsrAvailPlayers.Fields("pendingWaiver").Value = true then %>
								<strong>	<i class="fas fa-user-clock auctionText"></i></strong>
							<%end if%>
							<br><span class="gameTip"><%= objsrAvailPlayers.Fields("opponent")%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%= wtime%></small>
							</td>
							<td class="big" style="vertical-align:middle;text-align: center;width:12%;" class="text-center">
								<span class="badgeBlue"><%= round(objsrAvailPlayers.Fields ("barps").Value,0) %></span>
							</td>
							<td class="big" style="vertical-align:middle;text-align:center;width:12%;">
								<% if CDbl(objsrAvailPlayers.Fields("l5barps").Value) > CDbl(objsrAvailPlayers.Fields("barps").Value) then %>
									<span class="badgeUp"><%= round(objsrAvailPlayers.Fields ("l5barps").Value,0) %></span>
								<% elseif CDbl(objsrAvailPlayers.Fields("barps").Value) > CDbl(objsrAvailPlayers.Fields("l5barps").Value) then%>
									<span class="badgeDown"><%= round(objsrAvailPlayers.Fields ("l5barps").Value,0) %></span>
								<%else%>
									<span class="badgeEven"><%= round(objsrAvailPlayers.Fields ("l5barps").Value,0) %></span>
								<%end if %>														
							</td>
							<td class="big" style="vertical-align:middle;text-align: center;width:12%;" class="text-center">
								<% if IsNull(objsrAvailPlayers.Fields("usage").Value)  then %>
									<span class="badgeUsage">0</span>
								<%else%>
									<span class="badgeUsage"><%= round(objsrAvailPlayers.Fields("usage").Value,0) %></span>
								<%end if%>
							</td>